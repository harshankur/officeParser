/**
 * Shapes' names in DrawingML (the AST's `ShapeMetadata.geometry`: a preset name) and in ODF (a custom
 * shape's `draw:type`): which presets there are, and how one names the other's.
 *
 * @module utils/shapeNames
 */

/** DrawingML's preset shape names (ST_ShapeType): a `prstGeom` holds one of these and nothing else. */
export const DRAWINGML_PRESETS = new Set([
    'line', 'lineInv', 'triangle', 'rtTriangle', 'rect', 'diamond', 'parallelogram', 'trapezoid', 'nonIsoscelesTrapezoid', 'pentagon',
    'hexagon', 'heptagon', 'octagon', 'decagon', 'dodecagon', 'star4', 'star5', 'star6', 'star7', 'star8', 'star10', 'star12', 'star16',
    'star24', 'star32', 'roundRect', 'round1Rect', 'round2SameRect', 'round2DiagRect', 'snipRoundRect', 'snip1Rect', 'snip2SameRect',
    'snip2DiagRect', 'plaque', 'ellipse', 'teardrop', 'homePlate', 'chevron', 'pieWedge', 'pie', 'blockArc', 'donut', 'noSmoking',
    'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'stripedRightArrow', 'notchedRightArrow', 'bentUpArrow', 'leftRightArrow',
    'upDownArrow', 'leftUpArrow', 'leftRightUpArrow', 'quadArrow', 'leftArrowCallout', 'rightArrowCallout', 'upArrowCallout',
    'downArrowCallout', 'leftRightArrowCallout', 'upDownArrowCallout', 'quadArrowCallout', 'bentArrow', 'uturnArrow', 'circularArrow',
    'leftCircularArrow', 'leftRightCircularArrow', 'curvedRightArrow', 'curvedLeftArrow', 'curvedUpArrow', 'curvedDownArrow', 'swooshArrow',
    'cube', 'can', 'lightningBolt', 'heart', 'sun', 'moon', 'smileyFace', 'irregularSeal1', 'irregularSeal2', 'foldedCorner', 'bevel',
    'frame', 'halfFrame', 'corner', 'diagStripe', 'chord', 'arc', 'leftBracket', 'rightBracket', 'leftBrace', 'rightBrace', 'bracketPair',
    'bracePair', 'straightConnector1', 'bentConnector2', 'bentConnector3', 'bentConnector4', 'bentConnector5', 'curvedConnector2',
    'curvedConnector3', 'curvedConnector4', 'curvedConnector5', 'callout1', 'callout2', 'callout3', 'accentCallout1', 'accentCallout2',
    'accentCallout3', 'borderCallout1', 'borderCallout2', 'borderCallout3', 'accentBorderCallout1', 'accentBorderCallout2',
    'accentBorderCallout3', 'wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout', 'cloudCallout', 'cloud', 'ribbon', 'ribbon2',
    'ellipseRibbon', 'ellipseRibbon2', 'leftRightRibbon', 'verticalScroll', 'horizontalScroll', 'wave', 'doubleWave', 'plus',
    'flowChartProcess', 'flowChartDecision', 'flowChartInputOutput', 'flowChartPredefinedProcess', 'flowChartInternalStorage',
    'flowChartDocument', 'flowChartMultidocument', 'flowChartTerminator', 'flowChartPreparation', 'flowChartManualInput',
    'flowChartManualOperation', 'flowChartConnector', 'flowChartPunchedCard', 'flowChartPunchedTape', 'flowChartSummingJunction',
    'flowChartOr', 'flowChartCollate', 'flowChartSort', 'flowChartExtract', 'flowChartMerge', 'flowChartOfflineStorage',
    'flowChartOnlineStorage', 'flowChartMagneticTape', 'flowChartMagneticDisk', 'flowChartMagneticDrum', 'flowChartDisplay',
    'flowChartDelay', 'flowChartAlternateProcess', 'flowChartOffpageConnector', 'actionButtonBlank', 'actionButtonHome', 'actionButtonHelp',
    'actionButtonInformation', 'actionButtonForwardNext', 'actionButtonBackPrevious', 'actionButtonEnd', 'actionButtonBeginning',
    'actionButtonReturn', 'actionButtonDocument', 'actionButtonSound', 'actionButtonMovie', 'gear6', 'gear9', 'funnel', 'mathPlus',
    'mathMinus', 'mathMultiply', 'mathDivide', 'mathEqual', 'mathNotEqual', 'cornerTabs', 'squareTabs', 'plaqueTabs', 'chartX', 'chartStar',
    'chartPlus',
]);

/** ODF's names (draw:type) for the presets that have one of their own. */
export const PRESET_TO_ODF: Record<string, string> = {
    rect: 'rectangle', roundRect: 'round-rectangle', ellipse: 'ellipse', triangle: 'isosceles-triangle', rtTriangle: 'right-triangle',
    diamond: 'diamond', hexagon: 'hexagon', octagon: 'octagon', pentagon: 'pentagon', star4: 'star4', star5: 'star5', star8: 'star8',
    star24: 'star24', rightArrow: 'right-arrow', leftArrow: 'left-arrow', upArrow: 'up-arrow', downArrow: 'down-arrow',
    leftRightArrow: 'left-right-arrow', upDownArrow: 'up-down-arrow', parallelogram: 'parallelogram', trapezoid: 'trapezoid',
    chevron: 'chevron', homePlate: 'pentagon-right', plus: 'cross', cube: 'cube', can: 'can', heart: 'heart', sun: 'sun', moon: 'moon',
    smileyFace: 'smiley', lightningBolt: 'lightning', cloud: 'cloud', donut: 'ring', blockArc: 'block-arc', foldedCorner: 'paper',
    frame: 'frame', bevel: 'quad-bevel', wedgeRectCallout: 'rectangular-callout', wedgeRoundRectCallout: 'round-rectangular-callout',
    wedgeEllipseCallout: 'round-callout', cloudCallout: 'cloud-callout', verticalScroll: 'vertical-scroll', horizontalScroll: 'horizontal-scroll',
    flowChartProcess: 'flowchart-process', flowChartDecision: 'flowchart-decision', flowChartTerminator: 'flowchart-terminator',
    flowChartDocument: 'flowchart-document', flowChartConnector: 'flowchart-connector', flowChartInputOutput: 'flowchart-data',
};

const ODF_TO_PRESET: Record<string, string> = Object.fromEntries(Object.entries(PRESET_TO_ODF).map(([preset, odf]) => [odf, preset]));
Object.assign(ODF_TO_PRESET, { circle: 'ellipse', square: 'rect', 'round-square': 'roundRect', 'mso-spt100': 'rect' });

/** An ODF custom shape's `draw:type` as the preset it is, or undefined for one no preset is (a path of its own, `non-primitive`). */
export function presetOfOdfShape(type: string | null | undefined): string | undefined {
    if (!type) return undefined;
    return Object.prototype.hasOwnProperty.call(ODF_TO_PRESET, type) ? ODF_TO_PRESET[type] : DRAWINGML_PRESETS.has(type) ? type : undefined;
}
