# aoAdminNavigator

Admin Navigator is a Contensive addon that provides enhanced navigation and administrative tools for the Contensive platform.

## Contensive Patterns Reference

This project is built on the Contensive platform. Before implementing any Contensive-specific code (admin UI, addons, database models, collection XML, portals, page widgets, remote methods, etc.), you MUST read the relevant pattern documentation.

**Start here — read the patterns index to find the right pattern for your task:**
- [Contensive Patterns Index](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/index.md)

Then read the specific pattern(s) that apply to the work you are doing. Key patterns include:
- [AdminUI Pattern](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/adminui-pattern.md) — LayoutBuilderList, LayoutBuilderNameValue, filters, sorting, pagination, CSV export
- [Addon Collection Pattern](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/addon-collection-pattern.md) — collection XML, CDefs, Fields, packaging
- [Database Models Pattern](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/database-models-pattern.md) — typed C# model classes
- [Addon Pattern](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/addon-pattern.md) — addon architecture and types
- [Portal Pattern](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/portal-pattern.md) — portal-style admin interfaces
- [Page Widget Pattern](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/addon-page-widget-pattern.md) — design block widgets
- [Best Practices](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/best-practices-pattern.md) — error handling, try/catch, coding conventions
- [Security Best Practices](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/patterns/security-best-practices.md) — authentication, authorization, secure coding

Do NOT implement Contensive features from memory — always fetch and read the current pattern documentation first.

## GUIDs

When a GUID is needed (collection XML, addon definitions, CDefs, etc.), always generate a real unique GUID using the command line rather than hand-typing one:

```bash
python3 -c "import uuid; print(f'{{{uuid.uuid4()}}}'.upper())"
```

Never use patterned or sequential GUIDs like `{A1B2C3D4-...}` — they risk collisions.
