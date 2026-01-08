## 2025-05-15 - Add keyboard shortcut hints
**Learning:** Users expect common shortcuts like Ctrl+N and Ctrl+S to work globally or contextually, and tooltips claiming shortcuts exist without implementation is confusing.
**Action:** When adding tooltips with shortcut hints, ensure the bindings are actually implemented on the root window or relevant widget, and handle context switching (e.g., switch tab) if necessary.
