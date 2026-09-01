import { describe, expect, it } from "vitest";

import * as api from "../src/index";

describe("public api", () => {
    it("exports the sub-selection helper", () => {
        expect(typeof api.HtmlSubSelectionHelper.createHtmlSubselectionHelper).toBe("function");
    });

    it("exports the sub-selection attributes and selectors", () => {
        expect({
            DirectEditPlaceholderClassAndSelector: api.DirectEditPlaceholderClassAndSelector,
            DirectEditPlaceholderOutlineClassAndSelector: api.DirectEditPlaceholderOutlineClassAndSelector,
            FormatModeAttribute: api.FormatModeAttribute,
            HtmlSubSelectableClass: api.HtmlSubSelectableClass,
            HtmlSubSelectableSelector: api.HtmlSubSelectableSelector,
            SubSelectableAltObjectNameAttribute: api.SubSelectableAltObjectNameAttribute,
            SubSelectableDirectEdit: api.SubSelectableDirectEdit,
            SubSelectableDisplayNameAttribute: api.SubSelectableDisplayNameAttribute,
            SubSelectableHideOutlineAttribute: api.SubSelectableHideOutlineAttribute,
            SubSelectableObjectNameAttribute: api.SubSelectableObjectNameAttribute,
            SubSelectableRestrictingElementAttribute: api.SubSelectableRestrictingElementAttribute,
            SubSelectableSubSelectedAttribute: api.SubSelectableSubSelectedAttribute,
            SubSelectableTypeAttribute: api.SubSelectableTypeAttribute,
            SubSelectableUIAnchorAttribute: api.SubSelectableUIAnchorAttribute,
        }).toEqual({
            DirectEditPlaceholderClassAndSelector: { class: "direct-edit-placeholder", selector: ".direct-edit-placeholder" },
            DirectEditPlaceholderOutlineClassAndSelector: { class: "direct-edit-placeholder-outline", selector: ".direct-edit-placeholder-outline" },
            FormatModeAttribute: "format-mode",
            HtmlSubSelectableClass: "sub-selectable",
            HtmlSubSelectableSelector: ".sub-selectable",
            SubSelectableAltObjectNameAttribute: "data-sub-selection-alt-object-name",
            SubSelectableDirectEdit: "data-sub-selection-direct-edit",
            SubSelectableDisplayNameAttribute: "data-sub-selection-display-name",
            SubSelectableHideOutlineAttribute: "data-sub-selection-hide-outline",
            SubSelectableObjectNameAttribute: "data-sub-selection-object-name",
            SubSelectableRestrictingElementAttribute: "data-sub-selection-restricting-element",
            SubSelectableSubSelectedAttribute: "data-sub-selection-sub-selected",
            SubSelectableTypeAttribute: "data-sub-selection-type",
            SubSelectableUIAnchorAttribute: "data-sub-selection-ui-anchor",
        });
    });
});
