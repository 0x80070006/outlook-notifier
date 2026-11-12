using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;

namespace OutlookNotifier
{
    class MailInfo
    {
        public string Id;
        public string Sender;
        public string Subject;
        public string Preview;
        public DateTime Received;
    }

    class MeetingInfo
    {
        public string Subject;
        public string Location;
        public DateTime Start;
        public DateTime End;
    }

    class OutlookState
    {
        public bool Connected;
        public int Unread;
        public MailInfo LatestUnread;     // most recent unread message in the inbox, if any
        public MeetingInfo NextMeeting;   // ongoing or next meeting today, if any
        public int OtherMeetings;         // remaining meetings today besides NextMeeting

        public string Signature()
        {
            return Connected + "|" + Unread + "|" + (LatestUnread == null ? "" : LatestUnread.Id) + "|" +
                   (NextMeeting == null ? "" : NextMeeting.Subject + NextMeeting.Start.Ticks) + "|" + OtherMeetings;
        }
    }

    // Talks to classic Outlook (2016, 2019, 2021, Microsoft 365) through late-bound COM, so no
    // Office interop assembly is needed. All COM calls happen on one dedicated STA thread.
    class OutlookWatcher
    {
        const int PollMs = 3000;
        const int MeetingRefreshMs = 30000;
        const int ScanDepth = 25;
        const int olFolderInbox = 6;
        const int olFolderCalendar = 9;
        const int olMail = 43;
        const int olMinimized = 1;
        const int olNormalWindow = 2;
        const int olMeetingCanceled = 5;
        const int olMeetingReceivedAndCanceled = 7;
        const int RPC_E_CALL_REJECTED = unchecked((int)0x80010001);

        readonly BlockingCollection<Action> work = new BlockingCollection<Action>();
        readonly HashSet<string> seen = new HashSet<string>();
        readonly Queue<string> seenOrder = new Queue<string>();
        Thread thread;
        object outlook;
        bool primed;
        object mapi;              // Outlook namespace and inbox, kept between polls
        object inbox;
        string lastSignature;
        string lastFingerprint;   // item count | unread count | newest item, at the last scan
        MailInfo cachedLatest;
        MeetingInfo nextMeeting;
        int otherMeetings;
        DateTime nextMeetingRefresh = DateTime.MinValue;

        // Raised on the watcher thread. The flag is true when a new unread message just arrived.
        public event Action<OutlookState, bool> StateChanged;

        public void Start()
        {
            thread = new Thread(Loop) { IsBackground = true, Name = "Outlook" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        public void Stop()
        {
            if (!work.IsAddingCompleted) work.CompleteAdding();
        }

        public void OpenInbox() { Enqueue(() => OpenFolder(olFolderInbox, "outlook:inbox")); }
        public void OpenCalendar() { Enqueue(() => OpenFolder(olFolderCalendar, "outlook:calendar")); }

        void Enqueue(Action action)
        {
            if (!work.IsAddingCompleted) work.Add(action);
        }

        void Loop()
        {
            DateTime nextPoll = DateTime.UtcNow;
            while (true)
            {
                int wait = (int)Math.Max(0, (nextPoll - DateTime.UtcNow).TotalMilliseconds);
                Action action;
                try
                {
                    if (work.TryTake(out action, wait)) { Safe(action); continue; }
                }
                catch (InvalidOperationException) { break; }
                if (work.IsCompleted) break;
                Safe(Poll);
                FreeComWrappers();
                nextPoll = DateTime.UtcNow.AddMilliseconds(PollMs);
            }
            Disconnect();
        }

        // Every Outlook object reached through COM leaves a wrapper and a native proxy in this
        // process. They weigh almost nothing on the managed heap, so the GC would hardly ever run
        // on its own and the native memory kept growing (several GB after a day or two).
        static void FreeComWrappers()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        dynamic Outlook()
        {
            if (outlook != null) return outlook;
            if (!IsRunning("OUTLOOK")) return null;
            try { outlook = Marshal.GetActiveObject("Outlook.Application"); }
            catch { outlook = null; }
            return outlook;
        }

        public static bool IsRunning(string processName)
        {
            Process[] found = Process.GetProcessesByName(processName);
            foreach (var p in found) p.Dispose();
            return found.Length > 0;
        }

        void Disconnect()
        {
            Release(inbox);
            Release(mapi);
            inbox = null;
            mapi = null;
            outlook = null;
            lastFingerprint = null;
        }

        void Poll()
        {
            dynamic ol = Outlook();
            if (ol == null)
            {
                Publish(new OutlookState { Connected = false }, false, false);
                return;
            }
            if (mapi == null) mapi = ol.GetNamespace("MAPI");
            if (inbox == null) inbox = ((dynamic)mapi).GetDefaultFolder(olFolderInbox);
            dynamic box = inbox;

            // Cheap fingerprint first: the inbox is only scanned again when it actually changed.
            int unread = box.UnReadItemCount;
            dynamic items = box.Items;
            bool arrived = false;
            MailInfo latest = cachedLatest;
            try
            {
                int total = items.Count;
                items.Sort("[ReceivedTime]", true);
                string newestId = "";
                dynamic newest = items.GetFirst();
                if (newest != null)
                {
                    try { newestId = newest.EntryID; }
                    finally { Release(newest); }
                }
                string fingerprint = total + "|" + unread + "|" + newestId;
                if (fingerprint != lastFingerprint)
                {
                    lastFingerprint = fingerprint;
                    arrived = ScanForNewMail(items);
                    latest = unread > 0 ? FindLatestUnread(box) : null;
                    cachedLatest = latest;
                }
            }
            finally { Release(items); }

            bool meetingsRefreshed = false;
            if (arrived || DateTime.UtcNow >= nextMeetingRefresh)
            {
                RefreshMeetings(mapi);
                nextMeetingRefresh = DateTime.UtcNow.AddMilliseconds(MeetingRefreshMs);
                meetingsRefreshed = true;
            }

            var state = new OutlookState
            {
                Connected = true,
                Unread = unread,
                LatestUnread = latest,
                NextMeeting = nextMeeting,
                OtherMeetings = otherMeetings
            };
            Publish(state, arrived, meetingsRefreshed);
        }

        void Publish(OutlookState state, bool arrived, bool force)
        {
            string signature = state.Signature();
            if (!arrived && !force && signature == lastSignature) return;
            lastSignature = signature;
            if (StateChanged != null) StateChanged(state, arrived);
        }

        // True when an unread message appeared among the most recent inbox items since the last scan.
        // `items` is the inbox collection already sorted newest first.
        bool ScanForNewMail(dynamic items)
        {
            bool arrived = false;
            dynamic item = items.GetFirst();
            for (int i = 0; item != null && i < ScanDepth; i++)
            {
                try
                {
                    if ((int)item.Class == olMail)
                    {
                        string id = item.EntryID;
                        if (MarkSeen(id) && primed && (bool)item.UnRead) arrived = true;
                    }
                }
                finally { Release(item); }
                item = items.GetNext();
            }
            Release(item);
            primed = true;
            return arrived;
        }

        MailInfo FindLatestUnread(dynamic inbox)
        {
            dynamic all = inbox.Items;
            dynamic items = all.Restrict("[UnRead] = True");
            items.Sort("[ReceivedTime]", true);
            MailInfo result = null;
            dynamic item = items.GetFirst();
            for (int i = 0; item != null && i < 10 && result == null; i++)
            {
                try
                {
                    if ((int)item.Class == olMail)
                    {
                        string id = item.EntryID;
                        if (cachedLatest != null && cachedLatest.Id == id) result = cachedLatest;
                        else
                        {
                            result = Describe(item);
                            result.Id = id;
                        }
                    }
                }
                finally { Release(item); }
                if (result == null) item = items.GetNext();
            }
            Release(items);
            Release(all);
            return result;
        }

        static MailInfo Describe(dynamic item)
        {
            string body = "";
            try { body = Convert.ToString(item.Body) ?? ""; } catch { }
            if (body.Length > 600) body = body.Substring(0, 600);
            body = Regex.Replace(body, @"\s+", " ").Trim();

            string sender = Convert.ToString(item.SenderName);
            string subject = Convert.ToString(item.Subject);
            return new MailInfo
            {
                Sender = string.IsNullOrWhiteSpace(sender) ? null : sender.Trim(),
                Subject = string.IsNullOrWhiteSpace(subject) ? "(Sans objet)" : subject.Trim(),
                Preview = body,
                Received = item.ReceivedTime
            };
        }

        // Remaining meetings of the day (ongoing or upcoming), recurring occurrences included.
        void RefreshMeetings(dynamic ns)
        {
            MeetingInfo first = null;
            int total = 0;
            dynamic calendar = ns.GetDefaultFolder(olFolderCalendar);
            dynamic items = calendar.Items;
            items.Sort("[Start]");
            items.IncludeRecurrences = true;

            DateTime now = DateTime.Now;
            DateTime endOfDay = DateTime.Today.AddDays(1);
            string filter = "[Start] < '" + endOfDay.ToString("g", CultureInfo.CurrentCulture) +
                            "' AND [End] > '" + now.ToString("g", CultureInfo.CurrentCulture) + "'";
            dynamic found = items.Restrict(filter);
            dynamic item = found.GetFirst();
            for (int i = 0; item != null && i < 200; i++)
            {
                try
                {
                    int status = item.MeetingStatus;
                    if (!(bool)item.AllDayEvent && status != olMeetingCanceled && status != olMeetingReceivedAndCanceled)
                    {
                        DateTime start = item.Start;
                        DateTime end = item.End;
                        if (end > now && start < endOfDay)
                        {
                            total++;
                            if (first == null)
                            {
                                string subject = Convert.ToString(item.Subject);
                                first = new MeetingInfo
                                {
                                    Subject = string.IsNullOrWhiteSpace(subject) ? "(Sans titre)" : subject.Trim(),
                                    Location = (Convert.ToString(item.Location) ?? "").Trim(),
                                    Start = start,
                                    End = end
                                };
                            }
                        }
                    }
                }
                finally { Release(item); }
                item = found.GetNext();
            }
            Release(found);
            Release(items);
            Release(calendar);

            nextMeeting = first;
            otherMeetings = Math.Max(0, total - 1);
        }

        bool MarkSeen(string id)
        {
            if (string.IsNullOrEmpty(id) || !seen.Add(id)) return false;
            seenOrder.Enqueue(id);
            while (seenOrder.Count > 2000) seen.Remove(seenOrder.Dequeue());
            return true;
        }

        void OpenFolder(int folderId, string selectArg)
        {
            dynamic ol = Outlook();
            if (ol == null)
            {
                Process.Start(new ProcessStartInfo("outlook.exe", "/select " + selectArg) { UseShellExecute = true });
                return;
            }

            dynamic ns = ol.GetNamespace("MAPI");
            dynamic folder = ns.GetDefaultFolder(folderId);
            Release(ns);
            dynamic explorer = ol.ActiveExplorer();
            if (explorer == null)
            {
                folder.Display();
            }
            else
            {
                explorer.CurrentFolder = folder;
                if ((int)explorer.WindowState == olMinimized) explorer.WindowState = olNormalWindow;
                explorer.Activate();
                Release(explorer);
            }
            Release(folder);
            Native.BringProcessToFront("OUTLOOK");
        }

        void Safe(Action action)
        {
            try { action(); }
            catch (Exception e)
            {
                var com = e as COMException;
                if (com == null || com.ErrorCode != RPC_E_CALL_REJECTED)
                {
                    // Outlook closed or restarted: reconnect on the next poll.
                    Disconnect();
                    Log.Write(e.GetType().Name + ": " + e.Message);
                }
            }
        }

        static void Release(object o)
        {
            try { if (o != null && Marshal.IsComObject(o)) Marshal.ReleaseComObject(o); }
            catch { }
        }
    }

    // Sample content for the preview shown while the settings are open, and for --demo.
    static class Demo
    {
        public static OutlookState Fill(OutlookState real)
        {
            var s = new OutlookState
            {
                Connected = true,
                Unread = real == null ? 0 : real.Unread,
                LatestUnread = real == null ? null : real.LatestUnread,
                NextMeeting = real == null ? null : real.NextMeeting,
                OtherMeetings = real == null ? 0 : real.OtherMeetings
            };
            DateTime now = DateTime.Now;
            if (s.LatestUnread == null)
            {
                s.LatestUnread = new MailInfo
                {
                    Id = "demo",
                    Sender = "Camille Laurent",
                    Subject = "Point d'étape — projet Atlas",
                    Preview = "Bonjour, je te propose qu'on fasse le point jeudi à 10h pour valider la maquette et le planning de livraison.",
                    Received = now
                };
                s.Unread = Math.Max(s.Unread, 3);
            }
            if (s.NextMeeting == null)
            {
                DateTime start = now.Date.AddHours(now.Hour).AddMinutes(now.Minute < 30 ? 30 : 60);
                s.NextMeeting = new MeetingInfo { Subject = "Revue de sprint", Location = "Salle Atlas", Start = start, End = start.AddMinutes(45) };
                s.OtherMeetings = 6;
            }
            return s;
        }
    }
}
