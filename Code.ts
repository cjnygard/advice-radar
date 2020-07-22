/**
 * @license
 * Copyright 2019 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/** NOTE: the source code is Typescript, this is just the generated output.  *****/

// TODO
// use rayMap to cluster blips that are similar on the radar
// figure out how to reorder pages nicely, existing within a larger deck
// Make badge smaller.  Depends on scripting supporting modifying the padding of the badge element.
// Increase right-margin of blurb to avoid disappearing under badge.  Depends on scripting supporting modifying right-indent (or right-padding)
// Track merged content more granularly, so we can avoid blowing away additional slide content/formatting on update.

let FILTER_VALUE: string = 'solution';
let SUMMARY_SLIDE_LAYOUT_RANGE_NAME: string = 'radarSummarySlideLayoutId';
let RADAR_SLIDE_LAYOUT_RANGE_NAME: string = 'radarSlideLayoutId';
let APPENDIX_SLIDE_LAYOUT_RANGE_NAME: string = 'appendixSlideLayoutId';
let RECOMMENDATION_SLIDE_LAYOUT_RANGE_NAME: string = 'recommendationSlideLayoutId';
let PRESENTATION_TEMPLATE_RANGE_NAME: string = 'presentationTemplateId';
let PRESENTATION_RANGE_NAME: string = 'presentationId';
let HORIZON_CONFIG_RANGE_NAME: string = 'horizonConfig';
let CATEGORY_CONFIG_RANGE_NAME: string = 'categoryConfig';
let RADAR_CONFIG_RANGE_NAME: string = 'radarConfig';
let RADAR_ORDER_COLUMN: string = 'orderColumn';
let HORIZON_COLUMN: string = 'horizonColumn';
let DECK_CONFIG_RANGE_NAME: string = 'deckConfig';
let NODE_TABLE_RANGE_NAME: string = 'nodeTable';

// configuration for radar blurb blocks
let NUM_ROWS: number = 9;
let HEIGHT: number = 30;
let WIDTH: number = 130;
let LEFT_MARGIN: number = 20;
let TOP_MARGIN: number = 80;
let BADGE_WIDTH: number = 22;
let BADGE_HEIGHT: number = 12;
let TEXT_INDENT_END: number = 0.27;
let SPACING: number = 3;
let BLURB_COLOR: string = '#eeeeee';
let LOWER_RIGHT_HEIGHT: number = 375;
let LOWER_RIGHT_WIDTH: number = 700;
let BLURB_FONT_FAMILY: string = 'Open Sans';
let BLURB_FONT_WEIGHT: number = 100;
let BLURB_FONT_SIZE: number = 8;
let BADGE_FONT_FAMILY: string = 'Open Sans';
let BADGE_FONT_WEIGHT: number = 100;
let BADGE_FONT_SIZE: number = 5;
let BLIP_FONT_FAMILY: string = 'Open Sans';
let BLIP_FONT_WEIGHT: number = 500;
let BLIP_FONT_SIZE: number = 12;
let SELECTION_COLOR: string = '#0078bf';
let SELECTION_TEXT_COLOR: string = '#ffffff';
let ZONE_COLORS: string[] = ['#ee5ba0', '#0078bf', '#00bccd'];
let ZONE_RADIUS: number[] = [115, 210, 320];
let RADAR_COLOR: any = {};
let RADAR_RADIUS: any = {};
let RADIUS_VARIABILITY: number = 25;

interface KeyValue {
    key: any;
    value: any;
}

interface Savable {
    save(): void;
}
interface TableData<T extends KeyValueMap> extends Savable {
    data: T[];
}

interface Point {
    x: number;
    y: number;
}
interface KeyValueMap {
    [key: string]: any;
}

interface FormatMap<T> {
    [key: string]: (obj: T) => T;
}

interface Node extends KeyValueMap {
    dependency: string;
    type: string;
    subtype: string;
    category: string;
    classifier: string;
    visibility: number;
    maturity: number;
    horizon: string;
    impact: string;
    revenue: number;
    users: number;
    description: string;
    title: string;
    observations: string;
    evidence: string;
    implications: string;
    recommendations: string;
    outcomes: string;
    related: string;
    metrics: string;
    summaryOrder: string;
    radarPageId: string;
    radarEntryId: string;
    recommendationSlideId: string;
}

interface Edge extends KeyValueMap {
    parent: string;
    dependency: string;
    type: string;
    subtype: string;
    category: string;
    classifier: string;
    effort: string;
    impact: string;
    revenue: number;
    users: number;
    description: string;
}
interface DeckConfig extends KeyValueMap, Savable {
    presentationTemplateId: string;
    recommendationSlideLayoutId: string;
    radarSummarySlideLayoutId: string;
    radarSlideLayoutId: string;
    appendixSlideLayoutId: string;
    blankSlideLayoutId: string;
    newDeckName: string;
    presentationId: string;
    appendixSlideId: string;
    filterColumn: string;
    filterValue: string;
    groupColumn: string;
}

interface GroupConfig extends KeyValueMap {
    groupIdent: string;
    title: string;
    recommendationSlideLayoutId: string;
    radarAppendixId: string;
    radarPageId: string;
    radarSummaryPageId: string;
}

interface HorizonConfig extends KeyValueMap {
    horizonIdent: string;
    title: string;
}

interface CategoryConfig extends KeyValueMap {
    categoryIdent: string;
    title: string;
}

interface TestRestore extends KeyValueMap {
    key: string;
    value: number;
}

interface RayMapper {
    transform(node: Node): Point;
    index(): number;
}

interface DimensionMapper {
    [key: string]: number;
}

class SimpleRayMapper implements RayMapper {
    private count: number = 0;
    private increment: number = 17;
    private angleRange: number = 60;

    constructor(private radiusMap: DimensionMapper, private radiusKey: string) {}

    findAngle(node: Node): number {
        let angle: number = ((this.increment * this.count) % this.angleRange) + (90 - this.angleRange) / 2;
        return angle;
    }

    index(): number {
        return this.count;
    }

    transform(node: Node): Point {
        let angle: number = this.findAngle(node);
        let len: number = this.radiusMap[node[this.radiusKey]];
        let skewDist: number = Math.random() * RADIUS_VARIABILITY * 2 - RADIUS_VARIABILITY;
        let result: Point = fromPolar(angle, len + skewDist);
        ++this.count;
        return result;
    }
}

class ClusterRayMapper implements RayMapper {
    private angleMap: DimensionMapper = {};
    private count: number = 0;
    private increment: number = 17;
    private angleRange: number = 60;
    private minRadiusVariability: number = 10;
    constructor(private radiusMap: DimensionMapper, private radiusKey: string, private groupKey: string) {}

    findAngle(node: Node): number {
        let gk: string = node[this.groupKey];
        if (this.angleMap.hasOwnProperty(gk)) {
            return this.angleMap[gk];
        } else {
            let angle: number = ((this.increment * this.count) % this.angleRange) + (90 - this.angleRange) / 2;
            this.angleMap[gk] = angle;
            ++this.count;
            return angle;
        }
    }

    index(): number {
        return this.count;
    }

    transform(node: Node): Point {
        let angle: number = this.findAngle(node);
        let len: number = this.radiusMap[node[this.radiusKey]];
        let minRad: number = this.minRadiusVariability;
        let skewDist: number = Math.random() * (RADIUS_VARIABILITY - minRad) + minRad;
        let skewAngle: number = Math.random() * 360;
        let pos: Point = fromPolar(angle, len);
        let offset: Point = fromPolar(skewAngle, skewDist);
        let result: Point = { x: pos.x + offset.x, y: pos.y + offset.y };
        return result;
    }
}

// avoid passing these all around
let deckConfig: DeckConfig;
let groupConfig: TableData<GroupConfig>;
let horizonConfig: TableData<HorizonConfig>;
let categoryConfig: TableData<CategoryConfig>;

function fromPolar(angle: number, distance: number): Point {
    let rad: number = (angle * Math.PI) / 180.0;
    let x: number = Math.cos(rad) * distance;
    let y: number = Math.sin(rad) * distance;
    return { x: x, y: y };
}

function calculateCellRange_(row: string, col: number): string {
    return row + col + ':' + row + col;
}
function wrapTag_(text: string): string {
    return '{{' + text + '}}'; // TODO: change this to {{ and }}
}

function substituteValuesInSlide_(
    slide: GoogleAppsScript.Slides.Slide,
    kvList: KeyValue[]
): GoogleAppsScript.Slides.Slide {
    kvList.forEach(function (kv) {
        slide.replaceAllText(wrapTag_(kv.key), kv.value);
    });
    return slide;
}

function logSubstitutionData_(kvData: KeyValue[]) {
    kvData.forEach(function (kv) {
        Logger.log('[' + kv.key + ']:[' + kv.value + ']');
    });
}

function findLayoutSlide_(
    presentation: GoogleAppsScript.Slides.Presentation,
    layoutId: string
): GoogleAppsScript.Slides.Layout {
    let layout: GoogleAppsScript.Slides.Layout = (null as unknown) as GoogleAppsScript.Slides.Layout;
    presentation.getMasters().forEach(function (m: GoogleAppsScript.Slides.Master) {
        m.getLayouts().forEach(function (l: GoogleAppsScript.Slides.Layout) {
            if (l.getObjectId() == layoutId) {
                layout = l;
            }
        });
    });
    return layout;
}

function getKeyValueList_(kvMap: KeyValueMap): KeyValue[] {
    let result: KeyValue[] = [];
    for (var k in kvMap) {
        if (kvMap.hasOwnProperty(k)) {
            result.push({ key: k, value: kvMap[k] });
        }
    }
    return result;
}
function getArrayKeys_(values: any[], keyList: any[]) {
    let keys: any[] = null == keyList ? [] : [...keyList];

    for (var j = 1; j < values.length; j++) {
        keys.push(values[j][0]);
    }
    return keys;
}

function getArrayTags_(values: any[], prefix: string, kvList: KeyValue[]): any[] {
    let pairs: KeyValue[] = null == kvList ? [] : [...kvList];

    for (var j = 1; j < values.length; j++) {
        var row = values[j];
        pairs.push({ key: row[0], value: row[1] });
        pairs.push({ key: prefix + j, value: row[1] });
    }
    return pairs;
}

function findIndexOf_(entries: TableData<KeyValueMap>, keyColumnName: string, matchValue: string): number {
    let result: number = 0;
    for (var i = 0; i < entries.data.length; ++i) {
        if (entries.data[i][keyColumnName] == matchValue) {
            result = i + 1;
        }
    }
    return result;
}

function getTagList_(
    entries: TableData<KeyValueMap>,
    prefix: string,
    keyColumnName: string,
    dataColumnName: string,
    kvList?: KeyValue[]
): any[] {
    let pairs: KeyValue[] = null == kvList ? [] : [...kvList];
    let count: number = 1;
    entries.data.forEach(function (e: KeyValueMap) {
        pairs.push({ key: e[keyColumnName], value: e[dataColumnName] });
        pairs.push({ key: prefix + count, value: e[dataColumnName] });
        ++count;
    });
    return pairs;
}

function getDataTags_(obj: any, kvList?: KeyValue[], suffix?: string): KeyValue[] {
    let pairs: KeyValue[] = null == kvList ? [] : [...kvList];
    let tagSuffix: string = null == suffix ? '' : suffix;

    for (var i in obj) {
        if (obj.hasOwnProperty(i)) {
            pairs.push({ key: i + tagSuffix, value: obj[i] });
        }
    }
    return pairs;
}

function createMap_<T extends KeyValueMap>(keys: any[], values: any[], kvMap?: KeyValueMap): T {
    let kvPairs: KeyValueMap = null == kvMap ? new Map() : kvMap;
    for (var i = 0; i < keys.length && null != keys[i] && '' != keys[i]; ++i) {
        kvPairs[keys[i]] = values[i];
    }
    return kvPairs as T;
}

function parseTable_<T extends KeyValueMap>(data: any[][]): T[] {
    let headers: any[] = data[0];
    let entries: T[] = [];
    for (var i = 1; i < data.length; ++i) {
        entries.push(createMap_<T>(headers, data[i]));
    }
    return entries;
}

function restoreMap_(keys: any[], values: any[], kvMap: KeyValueMap) {
    let len: number = keys.length;
    for (var i = 0; i < len && null != keys[i] && '' != keys[i]; ++i) {
        if (kvMap.hasOwnProperty(keys[i])) {
            values[i] = kvMap[keys[i]];
        } else {
            values[i] = `missing [${i}]: [${keys[i]}]`;
            Logger.log(`${i}: missing value for [${keys[i]}]`);
        }
    }
}

function restoreTable_(data: any[], entries: KeyValueMap[]): KeyValueMap[] {
    let headers: any[] = data[0];
    let rows: number = Math.min(data.length, entries.length);
    for (let i: number = 0; i < rows; ++i) {
        restoreMap_(headers, data[i + 1], entries[i]);
    }
    return entries;
}

function parseObject_<T extends KeyValueMap>(data: any[][]): T {
    let obj: KeyValueMap = {};
    data.forEach(function (r: any[]) {
        obj[r[0]] = r[1];
    });
    return obj as T;
}

function restoreObject_(data: any[], obj: KeyValueMap) {
    data.forEach(function (r) {
        r[1] = obj[r[0]];
    });
}

function parseRangeObject_<T extends KeyValueMap>(
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    rangeName: string
): T {
    let values: any[][] = [];
    let range: GoogleAppsScript.Spreadsheet.Range | null = spreadsheet.getRangeByName(rangeName);
    if (null == range) {
        return {} as T;
    }
    values = range.getValues();
    let kvMap: KeyValueMap = parseObject_<T>(values);
    kvMap.save = function () {
        restoreObject_(values, this);
        range?.setValues(values);
    };
    return kvMap as T;
}

function parseRangeTable_<T extends KeyValueMap>(
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    rangeName: string
): TableData<T> {
    let obj: TableData<T> = new Object() as TableData<T>;
    let values: any[][] = [];
    let range: GoogleAppsScript.Spreadsheet.Range | null = spreadsheet.getRangeByName(rangeName);
    if (null == range) {
        return obj;
    }
    values = range.getValues();
    obj.data = parseTable_<T>(values);
    obj.save = function () {
        restoreTable_(values, obj.data);
        range?.setValues(values);
    };
    return obj;
}

function createRadarBlurb_(
    slide: GoogleAppsScript.Slides.Slide,
    recommendationSlide: GoogleAppsScript.Slides.Slide,
    horizon: any,
    kvList: KeyValue[],
    count: number
): GoogleAppsScript.Slides.Group {
    let index: number = count + 1;
    let blurb: GoogleAppsScript.Slides.Shape = slide.insertTextBox(wrapTag_('title'), 0, 0, WIDTH, HEIGHT);
    blurb.getFill().setSolidFill(BLURB_COLOR);
    blurb
        .getText()
        .getTextStyle()
        .setFontFamilyAndWeight(BLURB_FONT_FAMILY, BLURB_FONT_WEIGHT)
        .setFontSize(BLURB_FONT_SIZE);
    blurb.getText().getParagraphStyle().setIndentEnd(TEXT_INDENT_END);
    blurb.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    blurb.setLinkSlide(recommendationSlide);

    let badge: GoogleAppsScript.Slides.Shape = slide.insertTextBox(
        `${index}`,
        WIDTH - BADGE_WIDTH,
        0,
        BADGE_WIDTH,
        BADGE_HEIGHT
    );
    badge.getFill().setSolidFill(RADAR_COLOR[horizon]);
    badge
        .getText()
        .getTextStyle()
        .setFontFamilyAndWeight(BADGE_FONT_FAMILY, BADGE_FONT_WEIGHT)
        .setFontSize(BADGE_FONT_SIZE);
    badge
        .getText()
        .getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER)
        .setIndentEnd(0)
        .setIndentStart(0);
    badge.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    badge.setLinkSlide(recommendationSlide);

    let groupElements: GoogleAppsScript.Slides.PageElement[] = [];
    groupElements.push(slide.getPageElementById(blurb.getObjectId()));
    groupElements.push(slide.getPageElementById(badge.getObjectId()));

    let group: GoogleAppsScript.Slides.Group = slide.group(groupElements);
    group
        .setLeft(LEFT_MARGIN + (WIDTH + SPACING) * Math.floor(count / NUM_ROWS))
        .setTop(TOP_MARGIN + (count % NUM_ROWS) * (HEIGHT + SPACING));

    substituteValuesInSlide_(slide, kvList);

    return group;
}

function createRadarBlip_(
    slide: GoogleAppsScript.Slides.Slide,
    recommendationSlide: GoogleAppsScript.Slides.Slide,
    node: Node,
    rayMapper: RayMapper,
    kvMap: KeyValueMap,
    label: string
): GoogleAppsScript.Slides.Shape {
    let blip: GoogleAppsScript.Slides.Shape = slide.insertTextBox(label, 0, 0, 30, 30);
    Logger.log(rayMapper);
    let pos: Point = rayMapper.transform(node);
    blip.setLeft(LOWER_RIGHT_WIDTH - pos.x).setTop(LOWER_RIGHT_HEIGHT - pos.y);
    blip.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    blip.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    blip.getText()
        .getTextStyle()
        .setFontFamilyAndWeight(BLIP_FONT_FAMILY, BLIP_FONT_WEIGHT)
        .setFontSize(BLIP_FONT_SIZE);
    blip.setLinkSlide(recommendationSlide);

    return blip;
}

function appendNewSlide_(
    presentation: GoogleAppsScript.Slides.Presentation,
    layout: GoogleAppsScript.Slides.Layout,
    debugFlag: boolean
): GoogleAppsScript.Slides.Slide {
    if (null == layout) {
        return presentation.appendSlide();
    } else {
        return presentation.appendSlide(layout);
    }
}

function updateSlideLinks_(slide: GoogleAppsScript.Slides.Slide, slideMap: any) {
    slide
        .getPageElements()
        .filter(function (p: GoogleAppsScript.Slides.PageElement) {
            return p.getPageElementType() == SlidesApp.PageElementType.SHAPE;
        })
        .map(function (p: GoogleAppsScript.Slides.PageElement) {
            return p.asShape();
        })
        .filter(function (s: GoogleAppsScript.Slides.Shape) {
            return null != s.getLink();
        })
        .filter(function (s: GoogleAppsScript.Slides.Shape) {
            var link = s.getLink().getUrl();
            return slideMap.hasOwnProperty(link);
        })
        .forEach(function (s: GoogleAppsScript.Slides.Shape) {
            var link = s.getLink().getUrl();
            s.setLinkSlide(slideMap[link]);
        });
}

function updateSlideFromLayout_(
    presentation: GoogleAppsScript.Slides.Presentation,
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    layoutId: string,
    templateId: string,
    slideId: string,
    kvList: KeyValue[],
    reformatMap?: FormatMap<GoogleAppsScript.Slides.Shape>
) {
    Logger.log(`updateGroupSlide_: [${layoutId}]`);

    let slide: GoogleAppsScript.Slides.Slide;
    if (null == slideId || slideId == '' || null == (slide = presentation.getSlideById(slideId))) {
        let slideLayout: GoogleAppsScript.Slides.Layout = findLayoutSlide_(presentation, layoutId);
        slide = appendNewSlide_(presentation, slideLayout, true);
    } else {
        // Clear it out, to prepare for rebuilding from template
        slide.getPageElements().forEach(function (e: GoogleAppsScript.Slides.PageElement) {
            e.remove();
        });
    }

    // TODO:  Scan through layouts looking for the correct one
    var slideTemplate = findLayoutSlide_(presentation, templateId);
    if (null != slideTemplate) {
        let elements: GoogleAppsScript.Slides.PageElement[] = slideTemplate.getPageElements();
        elements.forEach(function (e: GoogleAppsScript.Slides.PageElement) {
            var newElement = slide.insertPageElement(e);
            // Check for match against a map of reformatting functions (for changing colors)
            if (null != reformatMap) {
                if (newElement.getPageElementType() == SlidesApp.PageElementType.SHAPE) {
                    var shape = newElement.asShape();
                    var text = shape.getText().asString().trim();
                    Logger.log(`Formatter: checking for [${text}]`);
                    if (reformatMap?.hasOwnProperty(text)) {
                        reformatMap?.[text]?.(shape);
                    }
                }
            }
        });
        substituteValuesInSlide_(slide, kvList);
    } else {
        Logger.log(`Cannot find slide template [${templateId}]`);
    }
    //  logSubstitutionData_(kvData);

    return slide;
}

function updateGroupSlide_(
    presentation: GoogleAppsScript.Slides.Presentation,
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    group: KeyValueMap,
    slideFieldName: string,
    templateId: string
): GoogleAppsScript.Slides.Slide {
    Logger.log(`updateGroupSlide_: [${templateId}] slideFieldName: ${slideFieldName}`);
    var slideId = group[slideFieldName];
    var substitutionData: KeyValue[] = getTagList_(horizonConfig, 'horizon', 'horizonIdent', 'title'); // Remap the horizon tag
    substitutionData = getDataTags_(group, substitutionData, '');

    let slide: GoogleAppsScript.Slides.Slide = updateSlideFromLayout_(
        presentation,
        spreadsheet,
        deckConfig.blankSlideLayoutId,
        templateId,
        slideId,
        substitutionData
    );
    group[slideFieldName] = slide.getObjectId();
    return slide;
}

function updateRadarSlide_(
    presentation: GoogleAppsScript.Slides.Presentation,
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    group: KeyValueMap,
    propertyName: string
): GoogleAppsScript.Slides.Slide {
    let slide: GoogleAppsScript.Slides.Slide = updateGroupSlide_(
        presentation,
        spreadsheet,
        group,
        propertyName,
        deckConfig[RADAR_SLIDE_LAYOUT_RANGE_NAME]
    );
    return slide;
}

function updateSummarySlide_(
    presentation: GoogleAppsScript.Slides.Presentation,
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    group: KeyValueMap,
    propertyName: string
): GoogleAppsScript.Slides.Slide {
    let slide: GoogleAppsScript.Slides.Slide = updateGroupSlide_(
        presentation,
        spreadsheet,
        group,
        propertyName,
        deckConfig[SUMMARY_SLIDE_LAYOUT_RANGE_NAME]
    );
    return slide;
}

function updateAppendixSlide_(
    presentation: GoogleAppsScript.Slides.Presentation,
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet
): GoogleAppsScript.Slides.Slide {
    let slideId: string = deckConfig.appendixSlideId;
    let substitutionData: KeyValue[] = getTagList_(horizonConfig, 'horizon', 'horizonIdent', 'title'); // Remap the horizon tags
    substitutionData = getTagList_(categoryConfig, 'category', 'categoryIdent', 'title', substitutionData); // Remap the category tag
    substitutionData = getTagList_(groupConfig, 'group', 'groupIdent', 'title', substitutionData); // getGroupTags_(spreadsheet);
    let slide: GoogleAppsScript.Slides.Slide = updateSlideFromLayout_(
        presentation,
        spreadsheet,
        deckConfig[APPENDIX_SLIDE_LAYOUT_RANGE_NAME],
        'noTemplate',
        slideId,
        substitutionData
    );
    deckConfig['appendixSlideId'] = slide.getObjectId();

    return slide;
}

function updateRecommendationSlide_(
    presentation: GoogleAppsScript.Slides.Presentation,
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    horizon: KeyValueMap,
    group: KeyValueMap,
    recommendation: Node
): GoogleAppsScript.Slides.Slide {
    Logger.log('updateSlide_: ');
    let slideId: string = recommendation.recommendationSlideId;

    let substitutionData: KeyValue[] = getTagList_(groupConfig, 'group', 'groupIdent', 'title'); // getGroupTags_(spreadsheet);
    substitutionData.push({ key: 'horizon', value: wrapTag_(recommendation.horizon) }); // Doing something tricky here to remap
    substitutionData.push({ key: 'category', value: wrapTag_(recommendation.category) }); // Doing something tricky here to remap
    substitutionData = getDataTags_(recommendation, substitutionData, '');
    substitutionData = getTagList_(horizonConfig, 'horizon', 'horizonIdent', 'title', substitutionData); // Remap the horizon tag
    substitutionData = getTagList_(categoryConfig, 'category', 'categoryIdent', 'title', substitutionData); // Remap the category tag
    let reformatMap: FormatMap<GoogleAppsScript.Slides.Shape> = {};
    reformatMap[wrapTag_('horizon')] = function (s: GoogleAppsScript.Slides.Shape): GoogleAppsScript.Slides.Shape {
        Logger.log(`Reformatted color for ${recommendation.title}`);
        s.getFill().setSolidFill(RADAR_COLOR[recommendation.horizon]);
        return s;
    };

    reformatMap[recommendation.impact] = function (s: GoogleAppsScript.Slides.Shape): GoogleAppsScript.Slides.Shape {
        Logger.log(`Reformatted Impact color for ${recommendation.title}`);
        s.getFill().setSolidFill(SELECTION_COLOR);
        s.getText().getTextStyle().setBold(true).setForegroundColor(SELECTION_TEXT_COLOR);
        return s;
    };

    let groupIdx: number = findIndexOf_(groupConfig, 'groupIdent', recommendation.subtype);
    Logger.log(`group[${recommendation.subtype}] idx[${groupIdx}]`);
    reformatMap[wrapTag_('group' + groupIdx)] = function (
        s: GoogleAppsScript.Slides.Shape
    ): GoogleAppsScript.Slides.Shape {
        Logger.log(`Reformatted Gropu color for ${recommendation.title}`);
        s.getFill().setSolidFill(SELECTION_COLOR);
        s.getText().getTextStyle().setBold(true).setForegroundColor(SELECTION_TEXT_COLOR);
        return s;
    };

    let slide: GoogleAppsScript.Slides.Slide = updateSlideFromLayout_(
        presentation,
        spreadsheet,
        deckConfig.blankSlideLayoutId,
        group[RECOMMENDATION_SLIDE_LAYOUT_RANGE_NAME],
        slideId,
        substitutionData,
        reformatMap
    );
    logSubstitutionData_(substitutionData);

    recommendation.recommendationSlideId = slide.getObjectId();

    return slide;
}

function testRestore() {
    let spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();

    let recs: TableData<TestRestore> = parseRangeTable_<TestRestore>(spreadsheet, 'testRestore');
    Logger.log(recs);
    Logger.log('Via map');
    Logger.log(recs.data[0]);
    recs.data[0]['value'] = 9;
    Logger.log(recs.data[0]);
    Logger.log('Via object');
    Logger.log(recs.data[1]);
    let rec: TestRestore = recs.data[1];
    rec.value = 8;
    recs.data[1] = rec;
    Logger.log(recs.data[1]);
    Logger.log(recs);
    recs.save();
}

function copyReportTemplate_(templateId: string, title: string): string {
    let date: string = Utilities.formatDate(new Date(), 'GMT+4', "yyyy-MM-dd'T'HH:mm");
    let deckTitle: string = title + '-' + date;

    let template: GoogleAppsScript.Drive.File = DriveApp.getFileById(templateId);
    let driveResponse: GoogleAppsScript.Drive.File = template.makeCopy(deckTitle);
    return driveResponse.getId();
}

function updateSlideDeck_(
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    presentation: GoogleAppsScript.Slides.Presentation
) {
    horizonConfig = parseRangeTable_<HorizonConfig>(spreadsheet, HORIZON_CONFIG_RANGE_NAME);
    categoryConfig = parseRangeTable_<CategoryConfig>(spreadsheet, CATEGORY_CONFIG_RANGE_NAME);
    groupConfig = parseRangeTable_<GroupConfig>(spreadsheet, RADAR_CONFIG_RANGE_NAME);
    let rowEntries: TableData<Node> = parseRangeTable_<Node>(spreadsheet, NODE_TABLE_RANGE_NAME);

    for (let i: number = 0; i < horizonConfig.data.length; ++i) {
        let hc: string = horizonConfig.data[i]['horizonIdent'];
        Logger.log(`[${i}] [${hc}]`);
        RADAR_RADIUS[hc] = ZONE_RADIUS[i];
        RADAR_COLOR[hc] = ZONE_COLORS[i];
    }

    let slideMap: any = {};

    let summaries: any[] = [];
    let radars: any[] = [];
    groupConfig.data.forEach(function (g: any) {
        g.summarySlide = updateSummarySlide_(presentation, spreadsheet, g, 'radarSummaryPageId');
        g.radarSlide = updateRadarSlide_(presentation, spreadsheet, g, 'radarPageId'); // In case we want to do something with it after update
        radars.push(g.radarSlide); // In case we want to do something with it after update
        summaries.push(g.summarySlide);
    });

    for (var j = 0; j < radars.length; ++j) {
        slideMap['#radar' + (j + 1)] = radars[j];
        slideMap['#summary' + (j + 1)] = summaries[j];
    }

    updateAppendixSlide_(presentation, spreadsheet);

    groupConfig.data.forEach(function (g: any) {
        Logger.log(`Processing group [${g.groupIdent}]`);
        // Logger.log(g);
        let groupFilter: string = g.groupIdent;
        g.radarIndexSlide = updateRadarSlide_(presentation, spreadsheet, g, 'radarAppendixId'); // In case we want to do something with it after update
        let count: number = 0;
        let summaryMap: any[] = [];
        let rayMapper: RayMapper = new SimpleRayMapper(RADAR_RADIUS, 'horizon');
        let rayMapperAppendix: RayMapper = new SimpleRayMapper(RADAR_RADIUS, 'horizon');

        horizonConfig.data.forEach(function (h: any) {
            let horizonFilter: string = h.horizonIdent;

            rowEntries.data.forEach(function (r: any) {
                Logger.log(`Processing row [${r.title}]`);
                let type: string = r[deckConfig.filterColumn];
                let group: string = r[deckConfig.groupColumn];
                let horizon: string = r.horizon;
                if (type == deckConfig.filterValue && group == groupFilter && horizon == horizonFilter) {
                    // Logger.log(`[${type}:${group}:${horizon}]`);

                    if (null != r.summaryOrder && r.summaryOrder > 0) {
                        summaryMap = getDataTags_(r, summaryMap, r.summaryOrder);
                    }

                    let kvMap: any = getDataTags_(r);

                    let recommendationSlide: GoogleAppsScript.Slides.Slide = updateRecommendationSlide_(
                        presentation,
                        spreadsheet,
                        h,
                        g,
                        r
                    );
                    updateSlideLinks_(recommendationSlide, slideMap);
                    let idx: number = rayMapper.index();
                    createRadarBlurb_(g.radarSlide, recommendationSlide, horizon, kvMap, idx);
                    createRadarBlip_(g.radarSlide, recommendationSlide, r, rayMapper, kvMap, `${1 + idx}`);
                    createRadarBlurb_(g.radarIndexSlide, recommendationSlide, horizon, kvMap, idx);
                    createRadarBlip_(g.radarIndexSlide, recommendationSlide, r, rayMapperAppendix, kvMap, `${1 + idx}`);

                    ++count;
                }
            });
        });
        substituteValuesInSlide_(g.summarySlide, summaryMap);
    });
    rowEntries.save();
    groupConfig.save();
}

function updateSlideDeck() {
    let spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();

    deckConfig = parseRangeObject_(spreadsheet, DECK_CONFIG_RANGE_NAME) as DeckConfig;
    Logger.log(deckConfig);

    let presentationId: string = deckConfig.presentationId; //presentationIdRange.getValue();
    let presentation: GoogleAppsScript.Slides.Presentation;
    if (null == presentationId || presentationId == '' || null == (presentation = SlidesApp.openById(presentationId))) {
        Logger.log(
            `Creating new deck [${deckConfig.newDeckName}] from template [${deckConfig.presentationTemplateId}]`
        );
        let presentationTemplateId: string = deckConfig.presentationTemplateId; //spreadsheet.getRangeByName(PRESENTATION_TEMPLATE_RANGE_NAME).getValue();
        presentationId = copyReportTemplate_(presentationTemplateId, deckConfig.newDeckName);
        deckConfig.presentationId = presentationId; //presentationIdRange.setValue(presentationId);
        presentation = SlidesApp.openById(presentationId);
    }
    updateSlideDeck_(spreadsheet, presentation);
    deckConfig.save();
}
