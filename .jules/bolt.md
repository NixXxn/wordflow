## 2024-05-22 - [Optimized Snippet Lookup]
**Learning:** Checking every possible suffix length in the key buffer (O(buffer_len)) is redundant when most lengths don't correspond to any snippet.
**Action:** Implemented a cache of unique snippet lengths in `SnippetManager`. The matching loop now iterates only over these specific lengths. This provided a ~6x speedup in the lookup logic (benchmark: 0.3s vs 2.1s for 100k ops).
