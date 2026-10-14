# Security Policy

## Supported versions

`pptx-automizer` is pre-1.0. Only the latest published `0.9.x` release
receives security fixes; older versions are not patched. Please upgrade to the
current release before reporting.

| Version        | Supported |
| -------------- | --------- |
| latest `0.9.x` | yes       |
| older          | no        |

## Reporting a vulnerability

Please report vulnerabilities through GitHub's private vulnerability
reporting: open the repository's **Security** tab and choose
**Report a vulnerability**. Please do not open a public issue for a security
problem.

A report is easiest to act on when it includes:

- the affected version of `pptx-automizer`
- a minimal reproduction, ideally a small sample `.pptx` and the code that
  triggers the problem
- the impact you see (what an attacker could achieve, and how)

Please reproduce the problem against this library before reporting. Generic
scanner findings or AI-generated reports without a reproduction cannot be
acted on, and a short note that you tried is always welcome.

## Scope

The library opens `.pptx` templates, which are ZIP archives of XML parts,
supplied by the calling application. In scope are issues such as:

- malicious template input (crafted ZIP or XML) that reaches runtime
- path traversal when extracting archives to disk
- XML parsing issues
- runtime dependencies with a vulnerability that is reachable through this
  library

Issues in `/website` (the docs-site build tooling) or in dev-only
dependencies are low priority.

## What to expect

This is a solo-maintained project. Reports are handled on a best-effort basis
without a fixed response time. A fix is published as a new release, with
credit in the `CHANGELOG.md` if you would like it.
