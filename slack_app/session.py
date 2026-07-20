"""In-memory session store keyed by the parent message's ts (== thread_ts
for replies in that thread). Good enough for a single always-running
process; a real database would only be needed for multi-instance/horizontal
scaling, which a small internal tool doesn't need."""

import threading


class Session:
    def __init__(self, channel, source="adimpact", title_override=None):
        self.channel = channel
        self.source = source  # "adimpact" or "adhawk"
        self.title_override = title_override
        self.spending_path = None
        self.spending_name = None
        self.creative_path = None
        self.creative_name = None

    def set_spending(self, path, name):
        self.spending_path = path
        self.spending_name = name

    def set_creative(self, path, name):
        self.creative_path = path
        self.creative_name = name


class SessionStore:
    def __init__(self):
        self._sessions = {}
        self._lock = threading.Lock()

    def create(self, thread_ts, channel, source="adimpact", title_override=None):
        with self._lock:
            self._sessions[thread_ts] = Session(channel, source=source, title_override=title_override)
        return self._sessions[thread_ts]

    def get(self, thread_ts):
        with self._lock:
            return self._sessions.get(thread_ts)

    def discard(self, thread_ts):
        with self._lock:
            self._sessions.pop(thread_ts, None)


sessions = SessionStore()
