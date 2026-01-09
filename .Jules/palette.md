## 2024-05-23 - Empty State in Tkinter Treeview
**Learning:** Tkinter's `ttk.Treeview` lacks a native "empty state" placeholder.
**Action:** Implement empty states by overlaying a centered `ttk.Label` using `place(relx=0.5, rely=0.5, anchor="center")` and toggling visibility based on item count.
