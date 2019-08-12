---
layout: default
title: vba-blocks
---

_A package manager for VBA_

Based heavily on the design of the much-lauded package manager for Rust, [Cargo](https://crates.io/),
vba-blocks aims to make building projects with VBA simple and powerful.
vba-blocks consists of three parts:

1. `vba-block.toml` is the manifest for a project, allowing you to specify dependencies, references, and source files for building your project
2. The developer add-in gets your project up-and-going and helps you install and update dependencies and source files
3. The vba-blocks CLI helps with starting your project and publishing packages for others to use

## Install

__Windows__

In powershell, run the following:

```shellsession
iwr https://vba-blocks.com/install.ps1 | iex
```

__Mac__

In terminal, run the following:

```shellsession
curl -fsSL https://vba-blocks.com/install.sh | sh
```

For more recent versions of Office for Mac, you will need to trust access to the VBA project object model for vba-blocks to work correctly:

<details>
  <summary>Trust access to the VBA project object model</summary>
  <ol>
    <li>Open Excel</li>
    <li>Click "Excel" in the menu bar</li>
    <li>Select "Preferences" in the menu</li>
    <li>Click "Security" in the Preferences dialog</li>
    <li>Check "Trust access to the VBA project object model" in the Security dialog</li>
 </ol>
 <a href="./images/trust-access-VBOM.png">Visual Instructions</a>
</details>


## vba-block.toml

The manifest format is based on Cargo's `Cargo.toml` and is used to specify the dependencies, references, and source files for your package or project.

For reference, here's what the `vba-block.toml` would look like for a relatively simple project like [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary):

```toml
[package]
name = "dictionary"
version = "1.4.1"
authors = ["Tim Hall <tim.hall.engr@gmail.com>"]

[src]
Dictionary = "Dictionary.cls"
```

An example of using the `dictionary` package in your project:

```toml
[project]
name = "My Project"
version = "0.0.0"
authors = ["tim.hall.engr@gmail.com"]

[src]
A = "A.bas"
B = "B.cls"
C = "C.frm"

[dependencies]
dictionary = "1.4.1"
```

For more details, see the [Manifest Format]({{baseurl}}/manifest/)

## CLI

```txt
Usage: vba-blocks [command] [options]

Commands:
  - new     Create a new project / package in a new directory
  - init     Initialize a new project / package in the current directory
  - build    Build project from manifest
  - export   Export src from built target
  - run      Run macro in document / add-in
```

Follow along on [GitHub](https://github.com/vba-blocks/vba-blocks)
