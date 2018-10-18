---
title: Manifest
---

# Manifest

The vba-blocks manifest (`vba-block.toml`) serves as the foundation for your project and provides information on your package, source, dependencies, references, and targets, as detailed below.

## `[package]` or `[project]`

`[package]` / `[project]` includes general information about your package.

```toml
[package]
name = "my_package"
version = "0.1.0"
authors = ["you@example.com"]
```

All three of the above fields (name, version, and authors) are mandatory for packages. Packages are designed to be included in other packages/projects, which projects are used for Office documents like Workbooks and have slightly different requirements.

```toml
[project]
name = "my_package"
target = "xlsm"
```

### `version`

vba-blocks is built on the concept of [Semantic Versioning](http://semver.org/), so the `version` field should follow these rules:

- Before you reach 1.0.0, anything goes.
- After 1.0.0, only make breaking changes when you increment the major version.
- After 1.0.0, don’t add any new public API (no new `Public` anything) in tiny versions. Always increment the minor version if you add any new `Public` enums, properties, methods, or anything else.
- Use version numbers with three numeric parts such as 1.0.0 rather than 1.0.

### `target`

`target` is used to define what application/extension to use when building your project. It can be a string for the extension, in which case `target/` includes the source files for creating the target. Otherwise, `type` and `path` can be used to define a custom target path

```toml
target = "xlsm"
# equivalent to target = { type = "xlsm", path = "target" }

target = { type = "xlam", path = "targets/xlam" }
```

## `[src]`

```toml
[src]
a = "src/a.bas"
b = "src/b.cls"
```

## `[dependencies]`

```toml
[dependencies]
dictionary = "^1"
json = { version = ">= 1.2, 1.5" }

[dependencies.utc]
version = "1.2.3"
```
