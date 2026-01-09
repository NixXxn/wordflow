## 2024-05-23 - [Weak Random Number Generation]
**Vulnerability:** Used `random.randint` for the `{random:min-max}` placeholder. The `random` module uses the Mersenne Twister algorithm, which is not cryptographically secure and is predictable.
**Learning:** Even in non-critical features like text expansion placeholders, users may rely on "random" features for security-sensitive contexts (e.g., password generation).
**Prevention:** Always use the `secrets` module (`secrets.SystemRandom`) for random number generation in any context where the output might be used for security purposes.
