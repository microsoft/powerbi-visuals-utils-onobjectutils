## 7.0.0
### Breaking changes
* the package is published as ES2020 modules instead of CommonJS, in line with the other `powerbi-visuals-utils-*` v7 packages
* `powerbi-visuals-api` moved to `devDependencies`: it is no longer installed transitively and has to be declared by the visual
* `HtmlSubSelectionHelper.getDataForElement` is now typed as `SubSelectionElementData | null`, matching what it returns for elements without the data attribute: consumers on `strictNullChecks` have to handle the `null` case
* `ISubSelectionHelper.getAllSubSelectables` takes a new optional `filterType` parameter: existing implementations outside this repo still compile, but they silently ignore the filter and return every sub-selectable until they are updated

### Fixed
* `createVisualSubSelectionForSingleObject` did not anchor the selection origin for `NumericText` and wrongly anchored it for `None`: the sub-selection type was matched with `in`, which checks array indices instead of values
* `getAllSubSelectables` ignored the filter when it was `SubSelectionStylesType.None` and returned every sub-selectable instead: the member is `0`, so it did not pass the truthiness check the filter was guarded with
* the sub-selection type of an element whose `data-sub-selection-type` attribute is missing or not a known member now falls back to `SubSelectionStylesType.None`, instead of putting `undefined` or `NaN` into the `subSelectionType` of the emitted sub-selection
* added the missing `typescript` dev dependency

### Other changes
* `isArrayEmpty` is now generic and returns a type predicate (`array is undefined | null | []`) instead of taking `any[]` and returning `boolean`: existing calls still compile, but the argument is no longer inferred as `any` and the result now narrows the array at the call site
* the published package is limited to `lib` and the docs via the `files` field, instead of relying on the ignore rules
* unit tests migrated to Vitest 4 and run in Chromium via `@vitest/browser` and Playwright, so the outline geometry is validated against real browser layout instead of a stubbed `getBoundingClientRect`
* ESLint migrated to the flat config, updated to ESLint 9 and typescript-eslint 8
* `strict` mode enabled
* new scripts: `prebuild`, `prepublishOnly`, `test:typecheck`, `lint:fix`
* CI runs audit, lint, typecheck, tests and build on Node 20 and 22; added Dependabot

## 6.0.2
* powerbi-visuals-api updated to 5.9.0
* added codeql.yml and build.yml