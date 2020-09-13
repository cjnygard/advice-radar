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
// Make badge smaller.  Depends on scripting supporting modifying the padding of the badge element.
// Increase right-margin of blurb to avoid disappearing under badge.  Depends on scripting supporting modifying right-indent (or right-padding)

let DECK_CONFIG_RANGE_NAME: string = 'configTable';

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
interface KeyTypeMap<T> {
    [key: string]: T;
}

interface FormatMap<T> {
    [key: string]: (obj: T) => T;
}

interface Config extends KeyValueMap, Savable {
    presentationTemplateId: string;
    presentationId: string;
    defaultLayoutId: string;
    newDeckName: string;

    slideTemplateRangeName: string;
    slideRangeName: string;
    valueMapRangeName: string;
    nodeTableRangeName: string;
    nodeSlidesRangeName: string;
    collectionRangeName: string;

    radarBlurbRows: number; // = 9,
    radarBlurbHeight: number; // = 30,
    radarBlurbWidth: number; // = 130,
    radarLeftMargin: number; // = 20,
    radarTopMargin: number; // = 80,
    radarBlurbBadgeWidth: number; // = 22,
    radarBlurbBadgeHeight: number; // = 12,
    radarBlurbTextIndentEnd: number; // = 0.27,
    radarBlurbSpacing: number; // = 3,
    radarBlurbColor: string; // = '#eeeeee',
    radarBlurbFontFamily: string; // = 'Open Sans',
    radarBlurbFontWeight: number; // = 100,
    radarBlurbFontSize: number; // = 8,
    radarBadgeFontFamily: string; // = 'Open Sans',
    radarBadgeFontWeight: number; // = 100,
    radarBadgeFontSize: number; // = 5,
    radarBlipFontFamily: string; // = 'Open Sans',
    radarBlipFontWeight: number; // = 500,
    radarBlipFontSize: number; // = 12,
    radarOriginHeight: number; // = 375,
    radarOriginWidth: number; // = 700,
    selectionColor: string; // = '#0078bf',
    selectionTextColor: string; // = '#ffffff',
    radarFieldName: string; // = 'horizon',
    zoneColors: string; // = '#00bccd,#0078bf,#ee5ba0',
    zoneRadius: string; // = "115,210,320",
    radarRadiusVariability: number; // = 25,
    radarZoneColors: string[];
    radarZoneRadius: number[];
    radarColor: KeyValueMap;
    radarRadius: KeyValueMap;
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
    blurb: string;
    observations: string;
    evidence: string;
    implications: string;
    recommendations: string;
    outcomes: string;
    related: string;
    metrics: string;
    summaryOrder: string;
    slideId: string;
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

interface ValueMapEntry extends KeyValueMap {
    field: string;
    ident: string;
    value: string;
}

interface ValueMap extends KeyValueMap {
    entries: TableData<ValueMapEntry>;
    fieldIdentMapper: KeyTypeMap<KeyValueMap>;
    uniqueKeys: string[];
    uniqueIdents: string[];
    indexMap: KeyValueMap;
    fieldValueEntries: KeyValue[];
}

interface ValueMapSpec extends KeyValueMap {
    identMap: KeyValueMap;
    fieldMap: KeyValueMap;
}

interface SlideTemplateConfig extends KeyValueMap {
    template: string;
    slideId: string;
}

interface SlideGeneratorSpec extends KeyValueMap {
    generator: string;
    title: string;
    collection: string;
    correlation: string;
    template: string;
    slideId: string;
    //    collectionMap: CollectionMap;
    slideStorageLocation: string;
    rayMapper: RayMapper;
}

interface NodeInfo extends KeyValueMap {
    node: Node;
    slide: KeyValueMap;
}

interface FilterSpec extends KeyValueMap {
    name: string;
    filter: string;
}

interface CriteriaSpec extends KeyValueMap {
    field: string;
    value: string;
}

interface Criteria extends KeyValueMap {
    spec: FilterSpec;
    criteriaSpec: CriteriaSpec[];
}

interface CriteriaMap extends KeyTypeMap<Criteria> {}

interface Collection extends KeyValueMap {
    criteria: Criteria;
    nodes: NodeInfo[];
}

interface CollectionMap extends KeyTypeMap<Collection> {}

interface SlideCache extends KeyTypeMap<KeyTypeMap<GoogleAppsScript.Slides.Slide>> {}

interface TypeCollectionIndexer extends KeyValueMap {
    linkMap: KeyTypeMap<KeyValueMap>; // maps #linkname (usuall #<type><index>) to type/collection names
    linkSlideMap: KeyValueMap; // maps #linkname directly to slide
    collectionLabelMap: KeyValueMap; // maps "collection<index>" to collection name  collection1 => quick-wins
    index: KeyValueMap;
    count: number;
}

interface Correlation extends KeyValueMap {
    collectionMap: CollectionMap;
    criteria: Criteria;
    indexer: TypeCollectionIndexer;
    cache: SlideCache;
    nodes: NodeInfo[];
}

interface CorrelationMap extends KeyTypeMap<Correlation> {}

interface TemplateSpec extends KeyValueMap {
    ident: string;
    templateId: string;
    layoutId: string;
}

interface TemplateMap extends KeyTypeMap<TemplateSpec> {}

interface RayMapper {
    transform(node: Node): Point;
    index(): number;
    next(): number;
}

// Syntactic sugar for a lack of actuall class implemntations
function init_<T>(func: (obj: T) => void): T {
    let result: T = new Object() as T;
    func(result);
    return result;
}

// Syntactic sugar for a lack of actuall class implemntations
function create_<T>(): T {
    let result: T = new Object() as T;
    return result;
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

    next(): number {
        return ++this.count;
    }

    transform(node: Node): Point {
        let angle: number = this.findAngle(node);
        let len: number = this.radiusMap[node[this.radiusKey]];
        let skewDist: number = Math.random() * config.radarRadiusVariability * 2 - config.radarRadiusVariability;
        let result: Point = fromPolar(angle, len + skewDist);
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

    next(): number {
        return this.count;
    }

    transform(node: Node): Point {
        let angle: number = this.findAngle(node);
        let len: number = this.radiusMap[node[this.radiusKey]];
        let minRad: number = this.minRadiusVariability;
        let skewDist: number = Math.random() * (config.radarRadiusVariability - minRad) + minRad;
        let skewAngle: number = Math.random() * 360;
        let pos: Point = fromPolar(angle, len);
        let offset: Point = fromPolar(skewAngle, skewDist);
        let result: Point = { x: pos.x + offset.x, y: pos.y + offset.y };
        return result;
    }
}

// avoid passing these all around
let config: Config;
let valueMap: ValueMap;
let slideTemplateTable: TemplateMap;
let slideGeneratorTable: TableData<SlideGeneratorSpec>;
let collectionMap: CollectionMap;
let correlationMap: CorrelationMap;
let criteriaMap: CriteriaMap;
let typeCollectionIndexer: TypeCollectionIndexer;

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

function createTemplateMap_(td: TableData<TemplateSpec>): TemplateMap {
    let result: TemplateMap = create_<TemplateMap>();
    td.data.forEach(function (t: TemplateSpec) {
        result[t.ident] = t;
    });
    return result;
}

function createCriteriaMap_(td: TableData<FilterSpec>): CriteriaMap {
    let result: CriteriaMap = create_<CriteriaMap>();
    // Logger.log('Processing filters');
    td.data.forEach(function (fs: FilterSpec) {
        let c: Criteria = init_<Criteria>(function (o: Criteria) {
            o.spec = fs;
            o.criteriaSpec = parseCriteriaSpec_(fs.filter);
        });
        // Logger.log(fs);
        // Logger.log(c);
        result[fs.ident] = c;
    });
    return result;
}

function createCollectionMap_(data: SlideGeneratorSpec[]): CollectionMap {
    let result: CollectionMap = create_<CollectionMap>();
    data.forEach(function (gen: SlideGeneratorSpec) {
        if (!result.hasOwnProperty(gen.collection)) {
            let c: Collection = init_<Collection>(function (o: Collection) {
                o.nodes = [];
                o.criteria = criteriaMap[gen.collection];
            });

            if (null == c.criteria) {
                Logger.log(`ERROR: criteria null for [${gen.collection}]`);
                Logger.log(criteriaMap);
            }
            result[gen.collection] = c;
        }
    });
    return result;
}

function createIndexer_(): TypeCollectionIndexer {
    let tci: TypeCollectionIndexer = init_<TypeCollectionIndexer>(function (o: TypeCollectionIndexer) {
        o.linkMap = create_<KeyTypeMap<KeyValueMap>>();
        o.linkSlideMap = create_<KeyValueMap>();
        o.collectionMap = create_<KeyValueMap>();
        o.index = create_<KeyValueMap>();
        o.count = 0;
    });
    return tci;
}

function createCorrelation_(correlation: string): Correlation {
    let result: Correlation = init_<Correlation>(function (o: Correlation) {
        o.collectionMap = create_<CollectionMap>();
        o.criteria = criteriaMap[correlation];
        o.cache = create_<SlideCache>();
        o.indexer = createIndexer_();
    });

    return result;
}

function createCorrelationMap_(td: TableData<SlideGeneratorSpec>): CorrelationMap {
    let result: CorrelationMap = create_<CorrelationMap>();
    let cache: KeyTypeMap<SlideGeneratorSpec[]> = create_<KeyTypeMap<SlideGeneratorSpec[]>>();
    let keys: string[] = [];
    td.data.forEach(function (gen: SlideGeneratorSpec) {
        if (!result.hasOwnProperty(gen.correlation)) {
            result[gen.correlation] = createCorrelation_(gen.correlation);
            cache[gen.correlation] = [];
            keys.push(gen.correlation);
        }
        cache[gen.correlation].push(gen);
    });
    keys.forEach(function (k: string) {
        result[k].collectionMap = createCollectionMap_(cache[k]);
    });

    return result;
}

function createValueMap_(td: TableData<ValueMapEntry>): ValueMap {
    let vm: ValueMap = new Object() as ValueMap;
    vm.fieldIdentMapper = new Object() as KeyTypeMap<KeyValueMap>;
    vm.fieldValueEntries = [];
    vm.uniqueKeys = [];
    vm.uniqueIdents = [];
    vm.indexMap = new Object() as KeyValueMap;
    let counts: KeyValueMap = new Object() as KeyValueMap;
    td.data.forEach(function (e: ValueMapEntry) {
        if (!counts.hasOwnProperty(e.field)) {
            counts[e.field] = 1;
            vm.uniqueKeys.push(e.field);
        }
        if (!vm.fieldIdentMapper.hasOwnProperty(e.field)) {
            vm.fieldIdentMapper[e.field] = new Object() as KeyValueMap;
        }
        vm.fieldIdentMapper[e.field][e.ident] = e.value;
        let key = `${e.field}${counts[e.field]}`;
        vm.indexMap[e.ident] = counts[e.field];
        vm.fieldValueEntries.push({ key: key, value: e.value });
        vm.fieldValueEntries.push({ key: e.ident, value: e.value });
        vm.uniqueIdents.push(e.ident);
        counts[e.field] = counts[e.field] + 1;
    });
    return vm;
}

function parseCriteriaSpec_(filter: string): CriteriaSpec[] {
    // format  <fieldName>:<matchValue>,<fieldName>:<matchValue>,....
    let result: CriteriaSpec[] = [];
    if (null != filter) {
        filter.split('|').forEach(function (f: string) {
            let pair: string[] = f.split(':');
            let fs: CriteriaSpec = init_<CriteriaSpec>(function (o: CriteriaSpec) {
                o.field = pair[0];
                o.value = pair[1];
            });
            result.push(fs);
        });
    }
    return result;
}

function filterMatch_(filter: CriteriaSpec, node: Node): boolean {
    return node[filter.field] == filter.value;
}

function filtersMatch_(filters: CriteriaSpec[], node: Node): boolean {
    let result: boolean = true;
    filters.forEach(function (f) {
        result = result && filterMatch_(f, node);
    });
    return result;
}

function createNodeInfo_(n: Node, s: KeyValueMap): NodeInfo {
    let result: NodeInfo = new Object() as NodeInfo;
    result.node = n;
    result.slide = s;
    return result;
}

function mergeNodeInfo_(rows: TableData<Node>, slides: TableData<KeyValueMap>): NodeInfo[] {
    let nodeInfo: NodeInfo[] = [];
    for (let j = 0; j < rows.data.length; ++j) {
        nodeInfo.push(createNodeInfo_(rows.data[j], slides.data[j]));
    }
    return nodeInfo;
}

function collectEntries_(map: CollectionMap, nodeInfo: NodeInfo[]): void {
    for (let i in map) {
        if (map.hasOwnProperty(i)) {
            map[i].nodes = nodeInfo.filter(function (n: NodeInfo): boolean {
                return filtersMatch_(map[i].criteria.criteriaSpec, n.node);
            });
        }
    }
}

function collectCorrelationEntries_(map: CorrelationMap, nodeInfo: NodeInfo[]): void {
    for (let i in map) {
        if (map.hasOwnProperty(i)) {
            map[i].nodes = nodeInfo.filter(function (n: NodeInfo): boolean {
                return filtersMatch_(map[i].criteria.criteriaSpec, n.node);
            });
            collectEntries_(map[i].collectionMap, map[i].nodes);
        }
    }
}

function getKeyValueList_(kvMap: KeyValueMap): KeyValue[] {
    let result: KeyValue[] = [];
    for (let k in kvMap) {
        if (kvMap.hasOwnProperty(k)) {
            result.push({ key: k, value: kvMap[k] });
        }
    }
    return result;
}
function getArrayKeys_(values: any[], keyList: any[]) {
    let keys: any[] = null == keyList ? [] : [...keyList];

    for (let j = 1; j < values.length; j++) {
        keys.push(values[j][0]);
    }
    return keys;
}

function getArrayTags_(values: any[], prefix: string, kvList: KeyValue[]): any[] {
    let pairs: KeyValue[] = null == kvList ? [] : [...kvList];

    for (let j = 1; j < values.length; j++) {
        let row = values[j];
        pairs.push({ key: row[0], value: row[1] });
        pairs.push({ key: prefix + j, value: row[1] });
    }
    return pairs;
}

function findIndexOf_(entries: TableData<KeyValueMap>, keyColumnName: string, matchValue: string): number {
    let result: number = 0;
    for (let i = 0; i < entries.data.length; ++i) {
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

    for (let i in obj) {
        if (obj.hasOwnProperty(i)) {
            pairs.push({ key: i + tagSuffix, value: obj[i] });
        }
    }
    return pairs;
}

function createMap_<T extends KeyValueMap>(keys: any[], values: any[], kvMap?: KeyValueMap): T {
    let kvPairs: KeyValueMap = null == kvMap ? new Map() : kvMap;
    for (let i = 0; i < keys.length && null != keys[i] && '' != keys[i]; ++i) {
        kvPairs[keys[i]] = values[i];
    }
    return kvPairs as T;
}

function parseTable_<T extends KeyValueMap>(data: any[][]): T[] {
    let headers: any[] = data[0];
    let entries: T[] = [];
    for (let i = 1; i < data.length; ++i) {
        entries.push(createMap_<T>(headers, data[i]));
    }
    return entries;
}

function restoreMap_(keys: any[], values: any[], kvMap: KeyValueMap) {
    let len: number = keys.length;
    for (let i = 0; i < len && null != keys[i] && '' != keys[i]; ++i) {
        if (kvMap.hasOwnProperty(keys[i])) {
            values[i] = kvMap[keys[i]];
        } else {
            values[i] = `missing [${i}]: [${keys[i]}]`;
            Logger.log(`ERROR: ${i}: missing value for [${keys[i]}]`);
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
    Logger.log(`Creating blurb, horizon [${horizon}]`);
    let blurb: GoogleAppsScript.Slides.Shape = slide.insertTextBox(
        wrapTag_('blurb'),
        0,
        0,
        config.radarBlurbWidth,
        config.radarBlurbHeight
    );
    blurb.getFill().setSolidFill(config.radarBlurbColor);
    blurb
        .getText()
        .getTextStyle()
        .setFontFamilyAndWeight(config.radarBlurbFontFamily, config.radarBlurbFontWeight)
        .setFontSize(config.radarBlurbFontSize);
    blurb.getText().getParagraphStyle().setIndentEnd(config.radarBlurbTextIndentEnd);
    blurb.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    blurb.setLinkSlide(recommendationSlide);

    let badge: GoogleAppsScript.Slides.Shape = slide.insertTextBox(
        `${index}`,
        config.radarBlurbWidth - config.radarBlurbBadgeWidth,
        0,
        config.radarBlurbBadgeWidth,
        config.radarBlurbBadgeHeight
    );
    badge.getFill().setSolidFill(config.radarColor[horizon]);
    badge
        .getText()
        .getTextStyle()
        .setFontFamilyAndWeight(config.radarBadgeFontFamily, config.radarBadgeFontWeight)
        .setFontSize(config.radarBadgeFontSize);
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
        .setLeft(
            config.radarLeftMargin +
                (config.radarBlurbWidth + config.radarBlurbSpacing) * Math.floor(count / config.radarBlurbRows)
        )
        .setTop(
            config.radarTopMargin +
                (count % config.radarBlurbRows) * (config.radarBlurbHeight + config.radarBlurbSpacing)
        );

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
    let pos: Point = rayMapper.transform(node);
    blip.setLeft(config.radarOriginWidth - pos.x).setTop(config.radarOriginHeight - pos.y);
    blip.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    blip.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    blip.getText()
        .getTextStyle()
        .setFontFamilyAndWeight(config.radarBlipFontFamily, config.radarBlipFontWeight)
        .setFontSize(config.radarBlipFontSize);
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
function registerSlide_(
    cache: SlideCache,
    tci: TypeCollectionIndexer,
    slide: GoogleAppsScript.Slides.Slide,
    generator: string,
    collection: string
): void {
    if (!cache.hasOwnProperty(collection)) {
        cache[collection] = new Object() as KeyTypeMap<GoogleAppsScript.Slides.Slide>;
    }
    cache[collection][generator] = slide;

    if (!tci.index.hasOwnProperty(collection)) {
        tci.count++;
        tci.index[collection] = tci.count;
        tci.collectionMap[`collection${tci.count}`] = collection;
    }

    let linkName: string = '#' + generator + tci.index[collection];
    Logger.log(`Register slide generator:collection: [${tci.count}] ${generator}:${collection} => ${linkName}`);

    tci.linkMap[linkName] = create_<KeyValueMap>();
    tci.linkMap[linkName].generator = generator;
    tci.linkMap[linkName].collection = collection;
    tci.linkSlideMap[linkName] = slide;
}

function updateSlideLinks_(slide: GoogleAppsScript.Slides.Slide, slideMap: any): void {
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
            let link = s.getLink().getUrl();
            return slideMap.hasOwnProperty(link);
        })
        .forEach(function (s: GoogleAppsScript.Slides.Shape) {
            let link = s.getLink().getUrl();
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
    Logger.log(`updateSlideFromLayout_: [${layoutId}]`);

    let slide: GoogleAppsScript.Slides.Slide;
    let slideLayout: GoogleAppsScript.Slides.Layout = findLayoutSlide_(presentation, layoutId);
    if (null == slideId || slideId == '' || null == (slide = presentation.getSlideById(slideId))) {
        slide = appendNewSlide_(presentation, slideLayout, true);
    } else {
        // Clear it out, to prepare for rebuilding from template
        slide.getPageElements().forEach(function (e: GoogleAppsScript.Slides.PageElement) {
            e.remove();
        });
    }

    // TODO:  Scan through layouts looking for the correct one
    let slideTemplate = findLayoutSlide_(presentation, templateId);
    if (null != slideTemplate) {
        let elements: GoogleAppsScript.Slides.PageElement[] = slideTemplate.getPageElements();
        elements.forEach(function (e: GoogleAppsScript.Slides.PageElement) {
            let newElement = slide.insertPageElement(e);
            // Check for match against a map of reformatting functions (for changing colors)
            if (null != reformatMap) {
                if (newElement.getPageElementType() == SlidesApp.PageElementType.SHAPE) {
                    let shape = newElement.asShape();
                    let text = shape.getText().asString().trim();
                    // Logger.log(`Formatter: checking for [${text}]`);
                    if (reformatMap?.hasOwnProperty(text)) {
                        reformatMap?.[text]?.(shape);
                    }
                }
            }
        });
        substituteValuesInSlide_(slide, kvList);
    } else {
        Logger.log(`ERROR: Cannot find slide template [${templateId}]`);
    }
    //  logSubstitutionData_(kvData);

    return slide;
}

function updateGroupSlide_(
    presentation: GoogleAppsScript.Slides.Presentation,
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    cache: SlideCache,
    tci: TypeCollectionIndexer,
    spec: SlideGeneratorSpec,
    entities: Collection
): GoogleAppsScript.Slides.Slide {
    let templateSpec: TemplateSpec = slideTemplateTable[spec.template];
    Logger.log(
        `updateGroupSlide_: [${templateSpec.templateId}] spec: ${spec.correlation}:${spec.collection}:${spec.generator}:${spec.title}`
    );
    let slideId = spec.slideId;
    let reformatMap: FormatMap<GoogleAppsScript.Slides.Shape> = {};

    var substitutionData: KeyValue[] = [...valueMap.fieldValueEntries];

    // create tag/values for summary entries
    entities?.nodes.forEach(function (ni: NodeInfo) {
        if (null != ni.node.summaryOrder && ni.node.summaryOrder != '') {
            substitutionData = getDataTags_(ni.node, substitutionData, ni.node.summaryOrder);
        }
    });
    substitutionData = getDataTags_(spec, substitutionData, '');

    let tag: string = `collection${valueMap.indexMap[spec.collection]}`;
    Logger.log(`reformatting for tag ${tag}`);
    reformatMap[wrapTag_(tag)] = function (s: GoogleAppsScript.Slides.Shape): GoogleAppsScript.Slides.Shape {
        s.getFill().setSolidFill(config.selectionColor);
        s.getText().getTextStyle().setBold(true).setForegroundColor(config.selectionTextColor);
        return s;
    };

    let slide: GoogleAppsScript.Slides.Slide = updateSlideFromLayout_(
        presentation,
        spreadsheet,
        'default' == templateSpec.layoutId ? config.defaultLayoutId : templateSpec.layoutId,
        templateSpec.templateId,
        slideId,
        substitutionData,
        reformatMap
    );
    spec.slideId = slide.getObjectId();
    // TODO  doublecheck this cache addressing
    registerSlide_(cache, tci, slide, spec.generator, spec.collection);
    return slide;
}

function updateEntitySlide_(
    presentation: GoogleAppsScript.Slides.Presentation,
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    cache: SlideCache,
    tci: TypeCollectionIndexer,
    spec: SlideGeneratorSpec,
    entities: Collection,
    nodeInfo: NodeInfo
): GoogleAppsScript.Slides.Slide {
    Logger.log(`updateEntitySlide_: [${nodeInfo.node.name}]`);
    let templateSpec: TemplateSpec = slideTemplateTable[spec.template];
    let slideId: string = nodeInfo.node[spec.slideStorageLocation];

    let substitutionData: KeyValue[] = [];
    substitutionData = getDataTags_(spec, substitutionData, '');

    valueMap.uniqueKeys.forEach(function (k: string) {
        substitutionData.push({ key: k, value: wrapTag_(nodeInfo.node[k]) });
    }); // Doing something tricky here to remap, this is order dependent in the substitution data

    substitutionData = getDataTags_(nodeInfo.node, substitutionData, '');
    substitutionData.push(...valueMap.fieldValueEntries);

    //logSubstitutionData_(substitutionData);

    let reformatMap: FormatMap<GoogleAppsScript.Slides.Shape> = {};
    reformatMap[wrapTag_(config.radarFieldName)] = function (
        s: GoogleAppsScript.Slides.Shape
    ): GoogleAppsScript.Slides.Shape {
        if (config.radarColor.hasOwnProperty(nodeInfo.node[config.radarFieldName])) {
            s.getFill().setSolidFill(config.radarColor[nodeInfo.node[config.radarFieldName]]);
        } else {
            // Logger.log(
            //    `Unable to reformat color for ${nodeInfo.node.name} no mapping for ${
            //        nodeInfo.node[config.radarFieldName]
            //    }`
            // );
        }
        return s;
    };

    reformatMap[nodeInfo.node.impact] = function (s: GoogleAppsScript.Slides.Shape): GoogleAppsScript.Slides.Shape {
        s.getFill().setSolidFill(config.selectionColor);
        s.getText().getTextStyle().setBold(true).setForegroundColor(config.selectionTextColor);
        return s;
    };

    valueMap.uniqueKeys.forEach(function (k: string) {
        let enumIdx: number = valueMap.indexMap[nodeInfo.node[k]];
        reformatMap[wrapTag_(k + enumIdx)] = function (
            s: GoogleAppsScript.Slides.Shape
        ): GoogleAppsScript.Slides.Shape {
            s.getFill().setSolidFill(config.selectionColor);
            s.getText().getTextStyle().setBold(true).setForegroundColor(config.selectionTextColor);
            return s;
        };
    });

    reformatMap[wrapTag_(`collection${tci.index[spec.collection]}`)] = function (
        s: GoogleAppsScript.Slides.Shape
    ): GoogleAppsScript.Slides.Shape {
        s.getFill().setSolidFill(config.selectionColor);
        s.getText().getTextStyle().setBold(true).setForegroundColor(config.selectionTextColor);
        return s;
    };

    let slide: GoogleAppsScript.Slides.Slide = updateSlideFromLayout_(
        presentation,
        spreadsheet,
        'default' == templateSpec.layoutId ? config.defaultLayoutId : templateSpec.layoutId,
        templateSpec.templateId,
        slideId,
        substitutionData,
        reformatMap
    );
    // logSubstitutionData_(substitutionData);

    updateSlideLinks_(slide, tci.linkSlideMap);
    let idx: number = spec.rayMapper.index();
    let blurbSlides: string[] = ['linked-radar', 'radar'];
    blurbSlides.forEach(function (s) {
        if (cache.hasOwnProperty(spec.collection) && cache[spec.collection].hasOwnProperty(s)) {
            let radarSlide = cache[spec.collection][s];
            if (null != radarSlide) {
                createRadarBlurb_(radarSlide, slide, nodeInfo.node[config.radarFieldName], substitutionData, idx);
                createRadarBlip_(radarSlide, slide, nodeInfo.node, spec.rayMapper, substitutionData, `${1 + idx}`);
            }
        }
    });
    spec.rayMapper.next();

    nodeInfo.slide[spec.slideStorageLocation] = slide.getObjectId();
    registerSlide_(cache, tci, slide, nodeInfo.node.name, spec.collection);

    return slide;
}

function copyReportTemplate_(templateId: string, title: string): string {
    let date: string = Utilities.formatDate(new Date(), 'GMT+4', "yyyy-MM-dd'T'HH:mm");
    let deckTitle: string = title + '-' + date;

    let template: GoogleAppsScript.Drive.File = DriveApp.getFileById(templateId);
    let driveResponse: GoogleAppsScript.Drive.File = template.makeCopy(deckTitle);
    return driveResponse.getId();
}

function updateSlide_(
    presentation: GoogleAppsScript.Slides.Presentation,
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    slideCache: SlideCache,
    tci: TypeCollectionIndexer,
    spec: SlideGeneratorSpec
): GoogleAppsScript.Slides.Slide {
    let slide: GoogleAppsScript.Slides.Slide;

    Logger.log('updateSlide_:');
    let cor: Correlation = correlationMap[spec.correlation];
    let col: CollectionMap = cor.collectionMap;
    switch (spec.generator) {
        case 'linked-radar':
        case 'radar':
        case 'summary':
        case 'static':
            slide = updateGroupSlide_(presentation, spreadsheet, cor.cache, cor.indexer, spec, col[spec.collection]);
            break;
        case 'linked-each':
        case 'each':
            col[spec.collection].nodes.forEach(function (ni: NodeInfo) {
                updateEntitySlide_(presentation, spreadsheet, cor.cache, cor.indexer, spec, col[spec.collection], ni);
            });
            break;
    }

    return slide;
}

function updateSlideDeck_(
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    presentation: GoogleAppsScript.Slides.Presentation
) {
    valueMap = createValueMap_(parseRangeTable_<ValueMapEntry>(spreadsheet, config.valueMapRangeName));
    slideTemplateTable = createTemplateMap_(parseRangeTable_<TemplateSpec>(spreadsheet, config.slideTemplateRangeName));
    slideGeneratorTable = parseRangeTable_<SlideGeneratorSpec>(spreadsheet, config.slideRangeName);
    //    collectionMap = createCollectionMap_(parseRangeTable_<FilterSpec>(spreadsheet, config.collectionRangeName));
    criteriaMap = createCriteriaMap_(parseRangeTable_<FilterSpec>(spreadsheet, config.collectionRangeName));
    let rowEntries: TableData<Node> = parseRangeTable_<Node>(spreadsheet, config.nodeTableRangeName);
    let rowSlideEntries: TableData<KeyValueMap> = parseRangeTable_<KeyValueMap>(
        spreadsheet,
        config.nodeSlidesRangeName
    );
    correlationMap = createCorrelationMap_(slideGeneratorTable);
    let nodes: NodeInfo[] = mergeNodeInfo_(rowEntries, rowSlideEntries);
    collectCorrelationEntries_(correlationMap, nodes);

    valueMap.uniqueIdents.forEach(function (s: string) {
        config.radarRadius[s] = config.radarZoneRadius[valueMap.indexMap[s] - 1];
        config.radarColor[s] = config.radarZoneColors[valueMap.indexMap[s] - 1];
    });

    let slideCache: SlideCache = new Object() as SlideCache;

    slideGeneratorTable.data.forEach(function (g: any) {
        if (g.slideId.includes('column')) {
            g.slideStorageLocation = g.slideId.split(':')[1];
        }
        g.rayMapper = new SimpleRayMapper(config.radarRadius, config.radarFieldName);
        updateSlide_(presentation, spreadsheet, slideCache, typeCollectionIndexer, g);
    });

    rowSlideEntries.save();
    slideGeneratorTable.save();
}

function updateSlideDeck() {
    let spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();

    config = parseRangeObject_(spreadsheet, DECK_CONFIG_RANGE_NAME) as Config;
    config.radarZoneColors = config.zoneColors.split(':');
    config.radarZoneRadius = config.zoneRadius.split('|').map(function (s: string): number {
        return +s;
    });
    config.radarRadius = new Object() as KeyValueMap;
    config.radarColor = new Object() as KeyValueMap;
    Logger.log(config);

    let presentationId: string = config.presentationId; //presentationIdRange.getValue();
    let presentation: GoogleAppsScript.Slides.Presentation;
    if (null == presentationId || presentationId == '' || null == (presentation = SlidesApp.openById(presentationId))) {
        Logger.log(`Creating new deck [${config.newDeckName}] from template [${config.presentationTemplateId}]`);
        let presentationTemplateId: string = config.presentationTemplateId; //spreadsheet.getRangeByName(PRESENTATION_TEMPLATE_RANGE_NAME).getValue();
        presentationId = copyReportTemplate_(presentationTemplateId, config.newDeckName);
        config.presentationId = presentationId; //presentationIdRange.setValue(presentationId);
        presentation = SlidesApp.openById(presentationId);
    }
    updateSlideDeck_(spreadsheet, presentation);
    config.save();
}
