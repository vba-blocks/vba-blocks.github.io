# vba-blocks

- Standard manifest (vba-block.toml)
- Add-in for Developer tab
- CLI for development

## Manifest

The block manifest (vba-block.toml) provides the foundation for building VBA projects and is based strongly on [Cargo's manifest format](http://doc.crates.io/manifest.html). 

VBA-Web example:

```toml
[package]
name = "VBA-Web"
version = "4.1.2"
authors = ["tim.hall.engr@gmail.com"]

description = "Connect VBA, Excel, Access, and Office for Windows and Mac to web services and the web"
license = "MIT"

[src]
WebClient = "src/WebClient.cls"
WebHelpers = "src/WebHelpers.bas"
WebRequest = "src/WebRequest.cls"
WebResponse = "src/WebResponse.cls"

[dev-src]
SpecAuthenticator = "specs/SpecAuthenticator.cls"
Specs_DigestAuthenticator = "specs/Specs_DigestAuthenticator.bas"
Specs_GoogleAuthenticator = "specs/Specs_GoogleAuthenticator.bas"
...

[features]
default = ["embed-dictionary", "json"]

# re-export embed/scripting features from VBA-Dictionary
embed-dictionary = ["VBA-Dictionary/embed"]
scripting-dictionary = ["VBA-Dictionary/scripting"]

json = { dependencies = ["VBA-JSON"] }
auto-proxy {}
async = {}
utf8 = {}

[dependencies]
VBA-Dictionary = "1.4.1"
VBA-JSON = {
  version = "2.2.0",
  optional = true
}

[dev-dependencies]
VBA-TDD = {
  version = "2.0.0-beta",
  features = ["workbook"]
}
VBA-DotEnv = "1.0.0"
VBA-LocalStorage = "1.0.0"
```

VBA-Dictionary example:

```toml
[package]
name = "VBA-Dictionary"
version = "1.4.1"
authors = ["tim.hall.engr@gmail.com"]

[src]
Dictionary = { path = "Dictionary.cls", optional = true }

[features]
embed = { src = ["Dictionary"] }
scripting = { references = ["Scripting"] }

[references]
Scripting = {
  description = "Microsoft Scripting Runtime"
  version = "1.0",
  guid = "{420B2830-E718-11CF-893D-00A0C9054228}",
  optional = true
}
```

Example usage:

```toml
[package]
name = "My Project"
version = "0.0.0"
publish = false

[dependencies]
VBA-Web = {
  version = "4.2.1",
  default-features = false,
  features = ["scripting-dictionary", "async"]
}
```

## Add-in

VBA-Blocks section in the Ribbon's Developer tab:

- Install - Install dependencies and dev-dependencies
- Check for updates
- Import - Import files based on src and dev-src
- Export - Export files based on src and dev-src
- Init - Initialize vba-block.toml
- Update - Update src and dev-src based on current

## CLI

```bash
vba-blocks --help

Usage: vba-blocks [options]
                  <cmd> [args]

Commands:

  init      Initialize block
  version   Update block version
  publish   Publish block

Options:

  -h, --help
  -V, --version
```

## References

- [VBA IDE CodeExport](https://github.com/spences10/VBA-IDE-Code-Export)
- [Cargo Manifest](http://doc.crates.io/manifest.html)
