## 2025-05-20 - Upgrade to Secure Random Generation
**Vulnerability:** Use of `random.randint` for the `{random:min-max}` placeholder. `random` is pseudo-random and predictable, unsuitable for security-sensitive contexts (e.g., generating passwords or PINs).
**Learning:** Even in non-security-focused apps, users might use random number generation for sensitive purposes. It's better to default to secure implementations when the cost is negligible.
**Prevention:** Use `secrets` module (specifically `secrets.randbelow`) for generating random numbers in user-facing features.
