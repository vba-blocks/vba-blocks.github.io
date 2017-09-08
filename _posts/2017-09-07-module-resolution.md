---
title: "Module Resolution"
description: "Deep dive into building module resolution for vba-blocks"
tags: [building-vba-blocks]
author: "Tim Hall"
---

One of the most important features of <b class="vba-blocks">vba-blocks</b> is package management: installing modules, classes, and references for packages from a variety of sources. Projects can be composed out of many small, focused libraries and it's super simple to share code across organizations and with the world. 

<b class="vba-blocks">vba-blocks</b> allows you to specify the direct dependencies that you want to use in your project and installs those direct dependencies, any transitive dependencies that the direct dependencies depend on, and so on until the entire dependency graph is resolved and installed. This is pretty standard behavior for package managers, but VBA has some fundamental properties that present an interesting challenge for package management:

1. All project code and dependencies share a global namespace
2. Dependencies may contain modules, classes, and/or references

These constraints need to be accounted for in the package manager and it falls on module resolution to sort them out. The module resolver has the following design requirements:

1. The dependency graph __must__ be flat
2. In evaluating dependencies, source names and references need to be considered

Looking to [Cargo](https://github.com/rust-lang/cargo) and [Yarn](https://github.com/yarnpkg/yarn) for inspiration, Yarn has a `--flat` install [option](https://yarnpkg.com/en/docs/cli/install#yarn-install---flat-a-classtoc-idtoc-yarn-install-flat-hreftoc-yarn-install-flata) that resolves the dependency tree using the following approach (referred to as _latest-solver_ here):

1. Resolve the entire dependency tree, selecting the latest version of each dependency that fulfills the parent's desired version range
2. Flatten the dependency tree, selecting from the versions of each package the latest that fulfills all of the version ranges for that package
3. If a version can't be selected for a dependency, the user is given control to explicitly override a dependency's version

This approach works well and is the first part of a 2-phase approach used by <b class="vba-blocks">vba-blocks</b>. The issue with the _latest-solver_ is that there may be a solution to the dependency graph that doesn't include the latest version of a dependency. For a more in-depth resolution approach, a constraint solver is used to search the _entire_ dependency graph for a solution. This can be a very intensive computation (technically [NP-Complete](https://research.swtch.com/version-sat)), so it's only used when the _latest-solver_ fails. <b class="vba-blocks">vba-blocks</b> uses a SAT solver ([logic-solver](https://github.com/meteor/logic-solver)) that makes it relatively straightforward to define the dependency tree and solve for the satisfying dependency versions. The solver has the following conditions:

__Direct Dependencies__

```js
// e.g. direct dependency "web: ^4.2.0"
const possibleVersions = ['web@4.2.0', 'web@4.2.1', 'etc'];
solver.require(Logic.exactlyOne(...possibleVersions));
```

__Transitive Dependencies__

```js
solver.require(Logic.atMostOne(...possibleVersions));
```

__Dependency Requirements__

```js
// e.g. module "web@4.2.0" depends on "dictionary: ^1"
const matching = getMatching(dictionary, '^1');
solver.require(Logic.implies('web@4.2.0'), Logic.or(...matching));
```

Finally, if both _latest-solver_ and _sat-solver_ fail to find a flat dependency graph, the user is prompted to manually override dependencies. With the flat tree requirement satisfied, <b class="vba-blocks">vba-blocks</b> then verifies that the selected dependency tree does not have conflicting source names or references. While this may be shifted into the selection process in the future, verifying requires more detailed information about each dependency than what is currently available during dependency resolution.

With that, a valid dependency tree has been determined and the build process can begin, which will be a topic of a future post. Once the build has completed successfully, a lockfile is written describing the dependency tree and project manifest. On re-builds, assuming nothing has changed in the manifest, the dependency tree is loaded from the lockfile and all calculations are skipped. Ideally, the only time module resolution occurs is when changing dependencies (via adding, removing, or updating) and for the majority of day-to-day usage the cached result can be used.

Inspired by best-in-class package managers like Cargo and Yarn, <b class="vba-blocks">vba-blocks</b> has a fast and thorough approach to module resolution that is designed specifically for the unique challenges of VBA. 