// powerbi-visuals-api only ships ambient `const enum` declarations, so the members have no
// runtime representation once the sources are transpiled by esbuild. This module provides the
// missing runtime values and is aliased in place of `powerbi-visuals-api` while testing.
// The `satisfies` clauses pin every copied value to the declared member, so `test:typecheck`
// fails if the API ever renumbers one instead of the drift only surfacing in production.
import type powerbiApi from "powerbi-visuals-api";

export const SubSelectionOutlineRestrictionType = {
    Clamp: 0,
    Clip: 1,
} as const satisfies {
    Clamp: powerbiApi.visuals.SubSelectionOutlineRestrictionType.Clamp;
    Clip: powerbiApi.visuals.SubSelectionOutlineRestrictionType.Clip;
};

export const SubSelectionStylesType = {
    None: 0,
    Text: 1,
    NumericText: 2,
    Shape: 3,
} as const satisfies {
    None: powerbiApi.visuals.SubSelectionStylesType.None;
    Text: powerbiApi.visuals.SubSelectionStylesType.Text;
    NumericText: powerbiApi.visuals.SubSelectionStylesType.NumericText;
    Shape: powerbiApi.visuals.SubSelectionStylesType.Shape;
};

export const SubSelectionOutlineType = {
    Group: 0,
    Rectangle: 1,
    Line: 2,
    Polygon: 3,
    Arc: 4,
} as const satisfies {
    Group: powerbiApi.visuals.SubSelectionOutlineType.Group;
    Rectangle: powerbiApi.visuals.SubSelectionOutlineType.Rectangle;
    Line: powerbiApi.visuals.SubSelectionOutlineType.Line;
    Polygon: powerbiApi.visuals.SubSelectionOutlineType.Polygon;
    Arc: powerbiApi.visuals.SubSelectionOutlineType.Arc;
};

export const SubSelectionOutlineVisibility = {
    None: 0,
    Hover: 1,
    Active: 2,
} as const satisfies {
    None: powerbiApi.visuals.SubSelectionOutlineVisibility.None;
    Hover: powerbiApi.visuals.SubSelectionOutlineVisibility.Hover;
    Active: powerbiApi.visuals.SubSelectionOutlineVisibility.Active;
};

const powerbi = {
    visuals: {
        SubSelectionOutlineRestrictionType,
        SubSelectionOutlineType,
        SubSelectionOutlineVisibility,
        SubSelectionStylesType,
    },
    extensibility: {},
};

export default powerbi;
