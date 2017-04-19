---
title: Manifest
---

# Manifest

The vba-blocks manifest (`vba-block.toml`) serves as the foundation for your project and provides information on your package, source, dependencies, references, and targets, as detailed below.

## `[package]`

`[package]` includes general information about your package.

```toml
[package]
name = "my_project"
version = "0.1.0"
authors = ["you@example.com"]
```

All three of the above fields are mandatory.

### `version`

vba-blocks is built on the concept of [Semantic Versioning](http://semver.org/), so the `version` field should follow these rules:

- Before you reach 1.0.0, anything goes.
- After 1.0.0, only make breaking changes when you increment the major version.
- After 1.0.0, don’t add any new public API (no new `Public` anything) in tiny versions. Always increment the minor version if you add any new `Public` enums, properties, methods, or anything else.
- Use version numbers with three numeric parts such as 1.0.0 rather than 1.0.

### `publish` (optional)

The `publish` field can be used to prevent a package from being published to vba-blocks by mistake.

```toml
[package]
# ...
publish = false
```

### `workspace` (optional)

The `workspace` field can be used to denote that the package is part of a workspace, allowing sharing of src, dependencies, or other information.

```toml
[package]
# ...
workspace = "path/to/root-workspace"
```

## `[src]`

```toml
[src]
a = "src/a.bas"
b = "src/b.cls"
c = "src/c.frm"
d = { path = "src/d.bas", optional = true }
```

## `[dependencies]`

```toml
[dependencies]
dictionary = "^1"
json = { version = ">= 1.2, 1.5", optional = true }

[dependencies.utc]
version = "1.2.3"
default-features = false
features = ["rfc3339"]
optional = true
```

## `[references]`

```toml
[references.Scripting]
version = "1.0"
guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
optional = true
```

## `[features]`

```toml
[features]
default = ["embed"]

emded = { src = ["Dictionary"] }
scripting = { references = ["Scripting"] }
```

## `[[targets]]`

```toml
[[targets]]
type = "xlsm"
path = "targets/xlsm"

[[targets]]
type = "xlam"
name = "my-addin"
path = "targets/xlam"
```
