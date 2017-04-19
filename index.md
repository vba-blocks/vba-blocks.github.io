---
layout: default
title: vba-blocks
---

# vba-blocks

_Coming soon: A package manager for VBA_

Based heavily on the design of the much-lauded package manager for Rust, [Cargo](https://crates.io/),
vba-blocks aims to make building projects with VBA simple and powerful.
vba-blocks consists of three parts:

1. `vba-block.toml` is the manifest for a project, allowing you to specify dependencies, references, and source files for building your project
2. The developer add-in gets your project up-and-going and helps you install and update dependencies and source files
3. The vba-blocks CLI helps with starting your project and publishing packages for others to use

## vba-block.toml

The manifest format is based on Cargo's `Cargo.toml` and is used to specify the dependencies, references, and source files for your package or project.

For reference, here's what the `vba-block.toml` would look like for a relatively simple project like [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary):

```toml
[package]
name = "dictionary"
version = "1.4.1"
author = ["Tim Hall <tim.hall.engr@gmail.com>"]

[src]
Dictionary = { path = "Dictionary.cls", optional = true }

[features]
default = ["embed"]

embed = { src = ["Dictionary"] }
scripting = { references = ["Scripting"] }

[dependencies]
# (none)

[references.Scripting]
version = "1.0"
guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
optional = true
```

An example of using the `dictionary` package in your project:

```toml
[package]
name = "My Project"
version = "0.0.0"
author = ["tim.hall.engr@gmail.com"]

# Explicitly prevent publishing of project / package
publish = false

[src]
A = "A.bas"
B = "B.cls"
C = "C.frm"

[dependencies]
dictionary = "1.4.1"
```

You can include features with dependencies by using the following expanded format:

```toml
# ...

# (remove the default "embed" feature and include "scripting")
[dependencies.dictionary]
version = "1.4.1"
default-features = false
features = ["scripting"]
```

For more details, see the [Manifest Format]({{baseurl}}/manifest/)

## Add-in

Add-ins for working with vba-blocks from Office.

- Dependencies: Add, remove, and update
- Source: Export on save and test

## CLI

```txt
Usage: vba-blocks [options]
                  <cmd> [args]

Commands:

  add       Add a block
  remove    Remove block(s)
  update    Update block(s) or all blocks
  build     Build project from dependencies and src
  test      Run project tests
  new       Create a new vba-blocks project
  init      Start a new vba-blocks project in the current directory
  version   Update block version
  publish   Publish block to vba-blocks

Options:

  -h, --help
  -V, --version
```

Follow along on [GitHub](https://github.com/vba-blocks/vba-blocks)
