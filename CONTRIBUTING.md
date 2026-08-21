# CONTRIBUTING.md

## Why This Matters

Fusion 360 is an amazing CAD/CAM system. It has a user-friendly interface, powerful modeling tools, simulation, and CNC programming features. Unfortunately, the built-in post-processors in the free version (especially for hobbyists and students) are often limited or don't work as expected.

Because of this, many people have to switch to other CAM systems, losing the simplicity and power of Fusion 360. I want to change that.

## Invitation to Contribute

**SmartPost** is a project created to give hobbyists full access to post-processing capabilities in Fusion 360, including:

- Tool changes without restrictions
- Removal of watermarks
- Support for any `.cps` post-processors
- **Rapid moves in toolpaths** (coming soon)

⚠️ I don't hack or patch Fusion. Everything is completely legal and fair — we use undocumented features that Autodesk has hidden for some reason.

---

## Calling All Python/CNC Developers

I'm building a tool to help Fusion 360 users stay license-compliant while getting the most from the software. As a solo developer, I could really use your help with:

Key Development Needs:

- Enhancing the `xml.cps` post-processor
- Improving G-code merging logic
- UI/UX refinements
- Better error handling/logging
- Multi-machine support
- Documentation improvements

How to Contribute:

1. Fork & modify the code (MIT licensed - just keep attribution)
2. Submit Pull Requests or share ideas via Issues
3. No strict rules - just keep changes focused

Before submitting a change, activate the project virtual environment and run:

```bash
pytest
ruff check .
ruff format --check .
python -m compileall .
pylint commands
git diff --check
```

### Testing downstream post compatibility

All selected `.cps` files process sections reconstructed by the Autodesk XML
importer. This is observably different from processing native CNC intermediate
data. A downstream post can therefore be correct but substantially slower when
called through SmartPost.

Pay particular attention to repeated calls to:

```javascript
section.getProperty(...)
section.getParameter(...)
```

Autodesk Fanuc revision 44236 is a confirmed regression: its section-scoped
lookups in `initializeSmoothing(_section)` take about 30 seconds on the XML test
fixture, while the older implementation using global `getProperty()` and
`getParameter()` completes in under one second.

Use `tmp/test-fanuc-performance.py` with an explicitly supplied installed
Autodesk `Applications/CAM360/post.exe` to reproduce the comparison. Never use
`tmp/postexe/post.exe`; that copy exists only for reverse-engineering artifacts.
Do not commit generated NC output or logs.

Do not patch, overwrite, or temporarily modify a user's selected `.cps` from
SmartPost. Fixes belong in the post maintained by its author or vendor. Until a
fixed post is available, document and recommend a known-compatible revision.

For Non-Coders:

- Report bugs/suggest features
- Help test new versions
- Star & share the project

**Every contribution helps make this tool better for Fusion 360 users!**

---

## Thank You

Every contribution matters. Even a small change to the [README](README.md) or code advice can impact hundreds of users. Thank you for reading!
