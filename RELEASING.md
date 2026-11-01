# Releasing

Push a version tag from `master`:

```
git tag -a vX.Y.Z -m "spoiwo X.Y.Z" && git push origin vX.Y.Z
```

CI then builds and tests both Scala versions, signs and publishes to Maven Central (`sbt ci-release`), and creates a GitHub Release with notes generated from the merged PRs. The version comes from the tag, so `build.sbt` never needs editing. Artifacts can take up to ~30 minutes to appear on Maven Central.

A published version can't be removed or replaced. A failed publish uploads nothing, so a broken release run can be fixed on `master` and the tag moved to the fix.

## Repository secrets

| Secret | Source |
|---|---|
| `SONATYPE_USERNAME` / `SONATYPE_PASSWORD` | User token from [central.sonatype.com](https://central.sonatype.com) → View Account → Generate User Token |
| `PGP_SECRET` | `gpg --armor --export-secret-keys <key-id> \| base64 \| gh secret set PGP_SECRET` |
| `PGP_PASSPHRASE` | Passphrase of that key |

The current signing key is `3165A793FF81B4DE` (fingerprint `39BC99BE5542A8475B812E013165A793FF81B4DE`, no expiry). Find a key's ID with `gpg --list-secret-keys --keyid-format long`.

- Sign in to the Central Portal with **email and password**. Signing in with Google gives a separate account without the `com.norbitltd` namespace, and its tokens can't publish.
- Run the `PGP_SECRET` export in a real terminal, not through a tool without a TTY: gpg needs to prompt for the passphrase, and otherwise an empty secret is stored silently. Check that `gpg --armor --export-secret-keys <key> | wc -c` is not 0 first.
- The signing key must stay published on `keyserver.ubuntu.com` so Central can verify signatures.
