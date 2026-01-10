## 2025-05-08 - Empty State Visuals
**Learning:** Users often feel lost when lists are empty without guidance. Adding a centered overlay with a call-to-action (like "Click 'New Snippet'") significantly improves onboarding and search feedback.
**Action:** When implementing list views (like Treeview), always include a hidden `ttk.Label` overlay that toggles visibility when the item count is zero. Use `place(relx=0.5, rely=0.5, anchor="center")` for perfect centering.
