# Security Policy

## Supported Versions

The following versions of MCPO-File-Generation-Tool are currently receiving security updates:

| Version | Supported          |
| ------- | ------------------ |
| 0.8.x   | :white_check_mark: |
| 0.9.x   | :white_check_mark: |
| < 0.8   | :x:                |

## Reporting a Vulnerability

We take security vulnerabilities seriously. If you discover a security issue, please follow these steps:

### Do NOT open a public GitHub issue

Public disclosure of security vulnerabilities can be exploited by malicious actors before a fix is available.

### Contact us privately

Please send a detailed report to:

**Email:** [contact@gliseman.tv](mailto:contact@gliseman.tv)

### What to include in your report

- Description of the vulnerability
- Steps to reproduce the issue (PoC preferred)
- Affected versions
- Severity assessment (if possible)
- Potential impact

### Timeline

- **Acknowledgment:** We will acknowledge receipt of your report within **48 hours**.
- **Assessment:** We will assess the vulnerability and provide an estimated timeline for a fix within **7 days**.
- **Fix & Disclosure:** We aim to release a fix within **30 days** of confirmation, depending on severity.
- **Public Disclosure:** After the fix is released, we will coordinate with you on public disclosure timing.

### Scope

The following areas are in scope for security reports:

- Path traversal and file injection vulnerabilities
- Remote Code Execution (RCE) through crafted files (DOCX, XLSX, PPTX, etc.)
- Authentication and authorization bypass
- Server-Side Request Forgery (SSRF)
- Denial of Service (DoS)
- Information disclosure
- Unsafe deserialization
- Template injection (e.g., Jinja2, Python string formatting)

### Out of scope

- Issues in third-party dependencies (please report these to the respective maintainers)
- Social engineering attacks against our team
- Issues that require physical access to a user's machine

## Bug Bounty

At this time, we do not offer a bug bounty program. Responsible disclosure is appreciated and we will credit reporters in our security advisories (unless they prefer to remain anonymous).

## Security Best Practices for Users

When deploying this tool, we recommend:

1. **Use HTTPS** for all communications with the MCP server.
2. **Set strong authentication** for your OpenWebUI instance (`JWT_SECRET`).
3. **Mount `FILE_EXPORT_DIR`** to a dedicated volume with limited permissions.
4. **Keep dependencies updated** — regularly run `pip install --upgrade -r requirements.txt`.
5. **Run as a non-root user** inside the Docker container.
6. **Restrict network access** to the MCP server and file export server.
7. **Use `PERSISTENT_FILES=false`** (default) so generated files are cleaned up automatically.

## Security Advisories

Security advisories will be published via GitHub Releases and the project's release notes.

## Acknowledgments

We thank the security researchers who responsibly disclose vulnerabilities to help keep this project safe for everyone.