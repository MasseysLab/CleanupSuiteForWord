# Free Open-Source Code Signing Plan

## Preferred Path

Apply to [SignPath Foundation](https://signpath.org/) for its sponsored
open-source code-signing program. The service and certificate are offered
without charge to accepted projects, but acceptance is discretionary.

CleanupSuite appears to meet the primary eligibility conditions:

- MIT is an OSI-approved open-source license;
- the source and build scripts are public;
- the project is actively maintained, released, and documented;
- it contains no hacking, advertising, telemetry, or unwanted-software features;
- installation is per user and provides repair and uninstall;
- document processing is local; and
- the repository has an automated Windows build-and-test workflow.

The self-contained Microsoft runtime is treated as a system/upstream component;
SignPath's final eligibility review determines whether the packaged runtime and
all other dependencies are acceptable.

## Prepared Policy Requirements

- [Code signing policy](../CODE_SIGNING_POLICY.md)
- [Privacy policy](../PRIVACY.md)
- Named project and signing roles
- Multi-factor-authentication requirement
- Explicit manual approval for each signing request
- Automated source, test, package, and security validation
- Safe unsigned-install instructions while an application is pending

## Work After Acceptance

1. Link the public GitHub repository and approved build workflow to SignPath.
2. Configure consistent product name and version metadata.
3. Build the self-contained engine and Setup from the linked source revision.
4. Have SignPath sign the engine before the installation manifest is generated.
5. Embed the signed engine, build the matched Setup, and have SignPath sign
   Setup.
6. Run the existing candidate manifest, signature, Defender, installer,
   production-manifest Word, Preview/Apply, Undo, and clean-machine gates.
7. Publish only the exact artifacts that passed those gates.

The Windows publisher displayed for sponsored signatures will be **SignPath
Foundation**, while CleanupSuite's product publisher remains **MasseysLab**.

## Fallback

If SignPath Foundation does not accept CleanupSuite, the project can remain free
and publish an explicitly unsigned release with hashes and the
[unsigned-install safety guide](Unsigned_Installation_Guide.md). The stronger
signed hybrid distribution contract should not be weakened until that rejection
and its exact implications are reviewed.

Estimated combined difficulty: **28–42%**.

Estimated professional desirability: **98–99.5%**.
