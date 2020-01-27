/*
 *  Power BI Visual CLI
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */
module powerbi.extensibility.visual {

    import valueFormatter = powerbi.extensibility.utils.formatting.valueFormatter;

    export module DataViewObjects {
        // Gets the value of the given object/property pair.
        export function getValue<T>(objects: DataViewObjects,
                                    propertyId: DataViewObjectPropertyIdentifier, defaultValue?: T): T {
            if (!objects) {
                return defaultValue;
            }
            const objectOrMap: DataViewObject = objects[propertyId.objectName];
            const object: DataViewObject = objectOrMap as DataViewObject;

            return DataViewObject.getValue(object, propertyId.propertyName, defaultValue);
        }

        // Gets an object from objects.
        export function getObject(objects: DataViewObjects,
                                  objectName: string, defaultValue?: DataViewObject): DataViewObject {
            if (objects && objects[objectName]) {

                return objects[objectName] as DataViewObject;
            } else {
                return defaultValue;
            }
        }

        // Gets a map of user-defined objects.
        export function getUserDefinedObjects(objects: DataViewObjects, objectName: string): DataViewObjectMap {
            if (objects && objects[objectName]) {

                return objects[objectName] as DataViewObjectMap;
            }
        }

        // Gets the solid color from a fill property.
        export function getFillColor(objects: DataViewObjects, propertyId: DataViewObjectPropertyIdentifier,
                                     defaultColor?: string): string {
            const value: Fill = getValue(objects, propertyId);
            if (!value || !value.solid) {
                return defaultColor;
            }

            return value.solid.color;
        }
    }

    export module DataViewObject {
        export function getValue<T>(object: DataViewObject, propertyName: string, defaultValue?: T): T {
            if (!object) {
                return defaultValue;
            }
            const propertyValue: T = object[propertyName] as T;
            if (propertyValue === undefined) {
                return defaultValue;
            }

            return propertyValue;
        }

        // Gets the solid color from a fill property using only a propertyName
        export function getFillColorByPropertyName(objects: DataViewObjects,
                                                   propertyName: string, defaultColor?: string): string {
            const value: Fill = DataViewObject.getValue(objects, propertyName);
            if (!value || !value.solid) {
                return defaultColor;
            }

            return value.solid.color;
        }
    }

    export interface ITextSettings {
        color: string;
        transparency: number;
        fontSize: number;
        alignment: string;
        alignmentV: string;
        direction: string;
        letterSpacing: number;
        lineHeight: number;
        wordSpacing: number;
        perspective: number;
        textIndent: number;
        lineIndent: number;
        textRotate: number;
        skewX: number;
        skewY: number;
    }

    export interface IStaticTextSettings {
        showColon: boolean;
        textPosition: string;
        textDecoration: string;
        textTransform: string;
        textShadow: string;
        textShadowBlur: string;
        textShadowColor: string;
        fontWeight: string;
        backgroundColor: string;
        transparency: number;
        fontFamily: string;
        boldStyle: boolean;
        underline: boolean;
        overline: boolean;
        strikeThrough: boolean;
        italicStyle: boolean;
        postText: string;
    }

    export interface IDynamicTextContainer {
        textContainer: string;
        lengthContainer: number;
    }

    export interface IDynamicTextSettings {
        backgroundColor: string;
        transparency: number;
        textDecoration: string;
        textTransform: string;
        textShadow: string;
        textShadowBlur: string;
        textShadowColor: string;
        fontWeight: string;
        fontFamily: string;
        boldStyle: boolean;
        underline: boolean;
        overline: boolean;
        strikeThrough: boolean;
        italicStyle: boolean;
    }
    export let questTextProperties: {
        textSettings: {
            color: DataViewObjectPropertyIdentifier;
            transparency: DataViewObjectPropertyIdentifier;
            fontSize: DataViewObjectPropertyIdentifier;
            postText: DataViewObjectPropertyIdentifier;
            alignment: DataViewObjectPropertyIdentifier;
            alignmentV: DataViewObjectPropertyIdentifier;
            direction: DataViewObjectPropertyIdentifier;
            letterSpacing: DataViewObjectPropertyIdentifier;
            lineHeight: DataViewObjectPropertyIdentifier;
            wordSpacing: DataViewObjectPropertyIdentifier;
            perspective: DataViewObjectPropertyIdentifier;
            textIndent: DataViewObjectPropertyIdentifier;
            lineIndent: DataViewObjectPropertyIdentifier;
            textRotate: DataViewObjectPropertyIdentifier;
            skewX: DataViewObjectPropertyIdentifier;
            skewY: DataViewObjectPropertyIdentifier;
        };
        staticTextSettings: {
            showColon: DataViewObjectPropertyIdentifier;
            textPosition: DataViewObjectPropertyIdentifier;
            textDecoration: DataViewObjectPropertyIdentifier;
            textTransform: DataViewObjectPropertyIdentifier;
            textShadow: DataViewObjectPropertyIdentifier;
            textShadowBlur: DataViewObjectPropertyIdentifier;
            textShadowColor: DataViewObjectPropertyIdentifier;
            fontWeight: DataViewObjectPropertyIdentifier;
            backgroundColor: DataViewObjectPropertyIdentifier;
            transparency: DataViewObjectPropertyIdentifier;
            postText: DataViewObjectPropertyIdentifier;
            fontFamily: DataViewObjectPropertyIdentifier;
            boldStyle: DataViewObjectPropertyIdentifier;
            italicStyle: DataViewObjectPropertyIdentifier;
            underline: DataViewObjectPropertyIdentifier;
            overline: DataViewObjectPropertyIdentifier;
            strikeThrough: DataViewObjectPropertyIdentifier;
        };
        dynamicSettings: {
            backgroundColor: DataViewObjectPropertyIdentifier;
            transparency: DataViewObjectPropertyIdentifier;
            textDecoration: DataViewObjectPropertyIdentifier;
            textTransform: DataViewObjectPropertyIdentifier;
            textShadow: DataViewObjectPropertyIdentifier;
            textShadowBlur: DataViewObjectPropertyIdentifier;
            textShadowColor: DataViewObjectPropertyIdentifier;
            fontWeight: DataViewObjectPropertyIdentifier;
            fontFamily: DataViewObjectPropertyIdentifier;
            boldStyle: DataViewObjectPropertyIdentifier;
            italicStyle: DataViewObjectPropertyIdentifier;
            underline: DataViewObjectPropertyIdentifier;
            overline: DataViewObjectPropertyIdentifier;
            strikeThrough: DataViewObjectPropertyIdentifier;
        }
    };

    questTextProperties = {
        dynamicSettings: {
            backgroundColor: {
                objectName: "Settings",
                propertyName: "backgroundColor"
            } as DataViewObjectPropertyIdentifier,
            boldStyle: { objectName: "Settings", propertyName: "boldStyle" } as DataViewObjectPropertyIdentifier,
            fontFamily: { objectName: "Settings", propertyName: "fontFamily" } as DataViewObjectPropertyIdentifier,
            fontWeight: { objectName: "Settings", propertyName: "fontWeight" } as DataViewObjectPropertyIdentifier,
            italicStyle: { objectName: "Settings", propertyName: "italicStyle" } as DataViewObjectPropertyIdentifier,
            overline: { objectName: "Settings", propertyName: "overline" } as DataViewObjectPropertyIdentifier,
            strikeThrough: {
                objectName: "Settings",
                propertyName: "strikeThrough"
            } as DataViewObjectPropertyIdentifier,
            textDecoration: {
                objectName: "Settings",
                propertyName: "textDecoration"
            } as DataViewObjectPropertyIdentifier,
            textShadow: { objectName: "Settings", propertyName: "textShadow" } as DataViewObjectPropertyIdentifier,
            textShadowBlur: {
                objectName: "Settings",
                propertyName: "textShadowBlur"
            } as DataViewObjectPropertyIdentifier,
            textShadowColor: {
                objectName: "Settings",
                propertyName: "textShadowColor"
            } as DataViewObjectPropertyIdentifier,
            textTransform: {
                objectName: "Settings",
                propertyName: "textTransform"
            } as DataViewObjectPropertyIdentifier,
            transparency: { objectName: "Settings", propertyName: "transparency" } as DataViewObjectPropertyIdentifier,
            underline: { objectName: "Settings", propertyName: "underline" } as DataViewObjectPropertyIdentifier
        },
        staticTextSettings: {
            backgroundColor: {
                objectName: "staticText",
                propertyName: "backgroundColor"
            } as DataViewObjectPropertyIdentifier,
            boldStyle: { objectName: "staticText", propertyName: "boldStyle" } as DataViewObjectPropertyIdentifier,
            fontFamily: { objectName: "staticText", propertyName: "fontFamily" } as DataViewObjectPropertyIdentifier,
            fontWeight: { objectName: "staticText", propertyName: "fontWeight" } as DataViewObjectPropertyIdentifier,
            italicStyle: { objectName: "staticText", propertyName: "italicStyle" } as DataViewObjectPropertyIdentifier,
            overline: { objectName: "staticText", propertyName: "overline" } as DataViewObjectPropertyIdentifier,
            postText: { objectName: "staticText", propertyName: "postText" } as DataViewObjectPropertyIdentifier,
            showColon: { objectName: "staticText", propertyName: "showColon" } as DataViewObjectPropertyIdentifier,
            strikeThrough: {
                objectName: "staticText",
                propertyName: "strikeThrough"
            } as DataViewObjectPropertyIdentifier,
            textDecoration: {
                objectName: "staticText",
                propertyName: "textDecoration"
            } as DataViewObjectPropertyIdentifier,
            textPosition: {
                objectName: "staticText",
                propertyName: "textPosition"
            } as DataViewObjectPropertyIdentifier,
            textShadow: { objectName: "staticText", propertyName: "textShadow" } as DataViewObjectPropertyIdentifier,
            textShadowBlur: {
                objectName: "staticText",
                propertyName: "textShadowBlur"
            } as DataViewObjectPropertyIdentifier,
            textShadowColor: {
                objectName: "staticText",
                propertyName: "textShadowColor"
            } as DataViewObjectPropertyIdentifier,
            textTransform: {
                objectName: "staticText",
                propertyName: "textTransform"
            } as DataViewObjectPropertyIdentifier,
            transparency: {
                objectName: "staticText",
                propertyName: "transparency"
            } as DataViewObjectPropertyIdentifier,
            underline: { objectName: "staticText", propertyName: "underline" } as DataViewObjectPropertyIdentifier
        },
        textSettings: {
            alignment: { objectName: "textSettings", propertyName: "alignment" } as DataViewObjectPropertyIdentifier,
            alignmentV: { objectName: "textSettings", propertyName: "alignmentV" } as DataViewObjectPropertyIdentifier,
            color: { objectName: "textSettings", propertyName: "color" } as DataViewObjectPropertyIdentifier,
            direction: { objectName: "textSettings", propertyName: "direction" } as DataViewObjectPropertyIdentifier,
            fontSize: { objectName: "textSettings", propertyName: "fontSize" } as DataViewObjectPropertyIdentifier,
            letterSpacing: {
                objectName: "textSettings",
                propertyName: "letterSpacing"
            } as DataViewObjectPropertyIdentifier,
            lineHeight: { objectName: "textSettings", propertyName: "lineHeight" } as DataViewObjectPropertyIdentifier,
            lineIndent: { objectName: "textSettings", propertyName: "lineIndent" } as DataViewObjectPropertyIdentifier,
            perspective: {
                objectName: "textSettings",
                propertyName: "perspective"
            } as DataViewObjectPropertyIdentifier,
            postText: { objectName: "textSettings", propertyName: "postText" } as DataViewObjectPropertyIdentifier,
            skewX: { objectName: "textSettings", propertyName: "skewX" } as DataViewObjectPropertyIdentifier,
            skewY: { objectName: "textSettings", propertyName: "skewY" } as DataViewObjectPropertyIdentifier,
            textIndent: { objectName: "textSettings", propertyName: "textIndent" } as DataViewObjectPropertyIdentifier,
            textRotate: { objectName: "textSettings", propertyName: "textRotate" } as DataViewObjectPropertyIdentifier,
            transparency: {
                objectName: "textSettings",
                propertyName: "transparency"
            } as DataViewObjectPropertyIdentifier,
            wordSpacing: { objectName: "textSettings", propertyName: "wordSpacing" } as DataViewObjectPropertyIdentifier
        },
    };

    export class TextEnhancer implements IVisual {
        private eventService: IVisualEventService;
        private target: d3.Selection<HTMLElement>;
        private dataViews: DataView;
        private staticTextSettings: IStaticTextSettings;
        private dynamicSettings: IDynamicTextSettings;
        private finalTextContainer: d3.Selection<HTMLElement>;
        private visualHost: IVisualHost;
        private transformed: string = "";
        private paddingType: string = "";
        private positionName: string = "";
        private positionVal: string = "";
        private degree360: number = 360;
        private degree180: number = 180;
        private degree90: number = 90;
        private degree270: number = 270;
        constructor(options: VisualConstructorOptions) {
            this.visualHost = options.host;
            this.eventService = options.host.eventService;
            this.target = d3.select(options.element);
            this.target.style({
                "cursor": "default",
                "overflow-y": "auto"
            });
        }

        public pointToPixel(pt: number): string {
            const pxPtRatio: number = 4 / 3;
            const pixelString: string = "px";

            return (pt * pxPtRatio) + pixelString;
        }

        public letSpace(ls: number): string {
            const lower = -3;
            const upper = 50;
            ls = ls < lower ? lower : ls > upper ? upper : ls;
            const pixelString: string = "px";

            return ls + pixelString;
        }

        public getLineHeight(lh: number): string {
            const defaultHeight = 1.6;
            lh = lh === (null || 0) ? defaultHeight : lh;
            const pixelString: string = "";

            return lh + pixelString;
        }

        public getWordSpace(ws: number): string {
            const lower = -3;
            const upper = 50;
            ws = ws === null ? 0 : (ws < lower ? lower : ws > upper ? upper : ws);
            const pixelString: string = "px";

            return ws + pixelString;
        }

        public getTextIndent(ti: number): string {
            const lower = -3;
            ti = ti == null ? 0 : (ti < lower ? lower : ti);
            const pixelString: string = "px";
            return ti + pixelString;
        }

        public getLineIndent(ti: number): string {
            const lower = -3;
            ti = ti == null ? 0 : (ti < lower ? lower : ti);
            const pixelString: string = "px";

            return ti + pixelString;
        }

        public getTextShadow(position: string, blur: string, color: string): string {
            let aShadowVar: number = 0;
            let bShadowVar: number = 0;
            let cShadowVar: number = 0;
            switch (position) {
                case "none":
                    return "";
                case "topLeft":
                    aShadowVar = -2;
                    bShadowVar = -2;
                    break;
                case "topCenter":
                    aShadowVar = 0;
                    bShadowVar = -2;
                    break;
                case "topRight":
                    aShadowVar = 2;
                    bShadowVar = -2;
                    break;
                case "middleLeft":
                    aShadowVar = -2;
                    bShadowVar = 0;
                    break;
                case "middleCenter":
                    aShadowVar = 0;
                    bShadowVar = 0;
                    break;
                case "middleRight":
                    aShadowVar = 2;
                    bShadowVar = 0;
                    break;
                case "bottomLeft":
                    aShadowVar = -2;
                    bShadowVar = 2;
                    break;
                case "bottomCenter":
                    aShadowVar = 0;
                    bShadowVar = 2;
                    break;
                case "bottomRight":
                    aShadowVar = 2;
                    bShadowVar = 2;
                    break;
                default: break;
            }
            switch (blur) {
                case "low":
                    cShadowVar = 2;
                    break;
                case "medium":
                    cShadowVar = 8;
                    break;
                case "high":
                    cShadowVar = 14;
                    break;
                default: break;
            }
            const pixelString: string = "px";

            return `${aShadowVar}${pixelString} ${bShadowVar}${pixelString} ${cShadowVar}${pixelString} ${color}`;
        }

        public getPerspective(fw: number): string {
            const pixelString: string = "px";

            return fw + pixelString;
        }

        public getSkew(sk: number): string {
            sk = sk === null ? 0 : (sk > this.degree360) ? this.degree360 : sk < 0 ? 0 : sk;
            const pixelString: string = "deg";

            return sk + pixelString;
        }

        public getSkewString(sx: number, sy: number): string {

            return `skewX(${this.getSkew(sx)}) skewY(${this.getSkew(sy)} )`;
        }

        public getDecimalPlacesCount(value: any): number {
            let decimalPlaces: number = 0;
            if (value > 0) {
                const arr: string[] = value.toString().split(".");
                if (!arr[1] && parseFloat(arr[1]) > 0) {
                    decimalPlaces = arr[1].length;
                }
            }

            return decimalPlaces;
        }

        public getDynamicTextValue(dataView: DataView): IDynamicTextContainer {
            let textValDynamicInput: any;
            let valueLength: number = 0;
            const categorical = dataView.categorical;
            if (dataView && categorical) {
                if (categorical.categories && categorical.categories[0] &&
                    categorical.categories[0].values) {
                    valueLength = categorical.categories[0].values.length;
                    textValDynamicInput = valueLength ? categorical.categories[0].values[0] : "(blank)";
                    if (categorical.categories[0].source && categorical.categories[0].source.format) {
                        const formatter: utils.formatting.IValueFormatter = valueFormatter.create({
                            format: categorical.categories[0].source.format
                        });
                        textValDynamicInput = formatter.format(textValDynamicInput);
                    }
                } else if (categorical.values && categorical.values[0] &&
                    categorical.values[0].values) {
                    valueLength = categorical.values[0].values.length;
                    textValDynamicInput = categorical.values[0].values[0] ?
                        categorical.values[0].values[0] : 0;
                    const upper = 4;
                    if (categorical.values[0] && categorical.values[0].source
                        && categorical.values[0].source.format) {
                        let decimalPlaces: number = this.getDecimalPlacesCount(textValDynamicInput);
                        decimalPlaces = decimalPlaces > upper ? upper : decimalPlaces;
                        const formatter: utils.formatting.IValueFormatter = valueFormatter.create({
                            format: categorical.values[0].source.format, precision: decimalPlaces, value: 1
                        });
                        textValDynamicInput = formatter.format(textValDynamicInput);
                    }
                }

                return {
                    lengthContainer: valueLength,
                    textContainer: textValDynamicInput
                };
            }
        }

        public getFontStyleClassName(settings: IDynamicTextSettings | IStaticTextSettings): string {
            let fontStyleClassName: string = "";
            if (settings.italicStyle) {
                fontStyleClassName = "tw_italic";
            }

            return fontStyleClassName;
        }

        public getTextDecoration(settings: IDynamicTextSettings | IStaticTextSettings): string {
            let textDecorationName: string = "";
            if (settings.underline) {
                textDecorationName += "underline ";
            }
            if (settings.overline) {
                textDecorationName += "overline ";
            }
            if (settings.strikeThrough) {
                textDecorationName += "line-through ";
            }

            return textDecorationName;
        }

        public convertHex(hex: string): string {
            hex = hex.replace("#", "");
            const rHexString: number = parseInt(hex.substring(0, 2), 16);
            const gHexString: number = parseInt(hex.substring(2, 4), 16);
            const bHexString: number = parseInt(hex.substring(4, 6), 16);

            return `rgb(${rHexString},${gHexString},${bHexString})`;
        }

        public getOpacityHex(transparency: number): string {
            const upper = 100;
            const max = 255;
            const lower = 6;
            const limit = 16;
            transparency = (upper - transparency);
            if (transparency === upper) {
                return "";
            } else {
                return transparency <= lower ? `0${Math.round((transparency / upper) * max).toString(limit).toUpperCase()}` :
                    Math.round((transparency / upper) * max).toString(limit).toUpperCase();
            }
        }

        public toRadians(angle: number): number {
            return angle * (Math.PI / this.degree180);
        }

        public getFontWeight(settings: IDynamicTextSettings | IStaticTextSettings): string {
            if (settings.boldStyle) {
                return "bold";
            } else {
                return "normal";
            }
        }

        public getTextTransform(settings: IDynamicTextSettings | IStaticTextSettings): string {
            return settings.textTransform;
        }
        public textSettingsTopValues(textSettings: ITextSettings, translate1: number,
                                     translate2: number, padding1: string, padding2: string,
                                     positionName: string, positionVal: number): void {
            const propVal: string = this.finalTextContainer.style("transform");
            this.transformed = `${propVal} translate(` + translate1 + `%, ` + translate2 + `%)`;
            this.paddingType = textSettings.lineIndent >= 0 ? padding1 : padding2;
            this.positionName = positionName;
            this.positionVal = positionVal + "%";
        }
        public textSettingsTop(textSettings: ITextSettings): void {
            const paddingVal: number = textSettings.lineIndent >= 0 ?
                textSettings.lineIndent : (-textSettings.lineIndent);
            switch (textSettings.direction) {
                case "horizontal-tb": {
                    this.textSettingsTopValues(textSettings, 0, 0, "padding-top", "padding-bottom", "top", 0);
                    if (textSettings.alignment === "right") {
                        this.finalTextContainer.style("float", "right");
                    } else if (textSettings.alignment === "center") {
                        this.finalTextContainer.style("width", "");
                    }
                }
                                      break;
                case "horizontal-bt": {
                    this.textSettingsTopValues(textSettings, 0, 0, "padding-bottom", "padding-top", "top", 0);
                    if (textSettings.alignment === "right") {
                        this.finalTextContainer.style("float", "right");
                    } else if (textSettings.alignment === "center") {
                        this.finalTextContainer.style("width", "");
                    }
                }
                                      break;
                case "vertical-rl": {
                    const lower = { left: -50, right: -100 };
                    const upper = { left: 50, right: 100 };
                    if (textSettings.alignment === "center") {
                        this.textSettingsTopValues(textSettings,
                            lower.left, 0, "padding-right", "padding-left", "left", upper.left);

                    } else if (textSettings.alignment === "right") {
                        this.textSettingsTopValues(textSettings,
                            lower.right, 0, "padding-right", "padding-left", "left", upper.right);
                    } else if (textSettings.alignment === "left") {
                        this.textSettingsTopValues(textSettings, 0, 0, "padding-right", "padding-left", "top", 0);
                    }
                }
                                    break;
                case "vertical-lr": {
                    const upper = 100;
                    const lower = 50;
                    if (textSettings.alignment === "center") {
                        this.textSettingsTopValues(textSettings, lower, 0,
                            "padding-left", "padding-right", "left", lower);
                    } else if (textSettings.alignment === "right") {
                        this.textSettingsTopValues(textSettings, upper, 0,
                            "padding-left", "padding-right", "left", upper);
                    } else if (textSettings.alignment === "left") {
                        this.textSettingsTopValues(textSettings, 0, 0, "padding-left", "padding-right", "top", 0);
                    }
                }
                                    break;
                default: break;
            }
            this.finalTextContainer = this.finalTextContainer
                .style("position", "relative")
                .style("transform", this.transformed)
                .style(this.positionName, this.positionVal)
                .style(this.paddingType, this.getLineIndent(paddingVal));
            this.finalTextContainer = this.finalTextContainer.append("div").classed("tw_pers", true);
        }
        public textSettingsMiddle(textSettings: ITextSettings): void {
            const paddingVal: number = textSettings.lineIndent >= 0 ?
                textSettings.lineIndent : (-textSettings.lineIndent);
            let transformed: string = "";
            let paddingType: string = "";
            let positionName: string = "";
            let positionVal: string = "";
            const propVal: string = this.finalTextContainer.style("transform");
            positionName = "top";
            positionVal = "50%";
            switch (textSettings.direction) {
                case "horizontal-tb": {
                    transformed = `${propVal}  translate(0%, -50%)`;
                    paddingType = textSettings.lineIndent >= 0 ?
                        "padding-top" : "padding-bottom";

                    if (textSettings.alignment === "right") {
                        this.finalTextContainer.style("float", "right");
                    } else if (textSettings.alignment === "center") {
                        this.finalTextContainer.style("width", "");
                    }
                }
                                      break;
                case "horizontal-bt": {
                    transformed = `${propVal}  translate(0%, 50%)`;
                    paddingType = textSettings.lineIndent >= 0 ?
                        "padding-bottom" : "padding-top";

                    if (textSettings.alignment === "right") {
                        this.finalTextContainer.style("float", "right");
                    } else if (textSettings.alignment === "center") {
                        this.finalTextContainer.style("width", "");
                    }
                }
                                      break;
                case "vertical-rl": {
                    if (textSettings.alignment === "center") {
                        transformed = `${propVal}  translate(-50%, -50%)`;
                        paddingType = textSettings.lineIndent >= 0 ?
                            "padding-right" : "padding-left";
                        this.finalTextContainer = this.finalTextContainer
                            .style("left", "50%");
                    } else if (textSettings.alignment === "right") {
                        transformed = `${propVal}  translate(-100%, -50%)`;
                        paddingType = textSettings.lineIndent >= 0 ?
                            "padding-right" : "padding-left";
                        this.finalTextContainer = this.finalTextContainer
                            .style("left", "100%");
                    } else if (textSettings.alignment === "left") {
                        transformed = `${propVal}  translate(0%, -50%)`;
                        paddingType = textSettings.lineIndent >= 0 ?
                            "padding-right" : "padding-left";
                    }
                }
                                    break;
                case "vertical-lr": {
                    if (textSettings.alignment === "center") {
                        transformed = `${propVal}  translate(50%, 50%)`;
                        paddingType = textSettings.lineIndent >= 0 ?
                            "padding-left" : "padding-right";
                        this.finalTextContainer = this.finalTextContainer
                            .style("left", "50%");
                    } else if (textSettings.alignment === "right") {
                        transformed = `${propVal}  translate(100%, 50%)`;
                        paddingType = textSettings.lineIndent >= 0 ?
                            "padding-left" : "padding-right";
                        this.finalTextContainer = this.finalTextContainer
                            .style("left", "100%");
                    } else if (textSettings.alignment === "left") {
                        transformed = `${propVal}  translate(0%, 50%)`;
                        paddingType = textSettings.lineIndent >= 0 ?
                            "padding-left" : "padding-right";
                    }
                }
                                    break;
                default: break;
            }
            this.finalTextContainer = this.finalTextContainer
                .style("position", "relative")
                .style("transform", transformed)
                .style(positionName, positionVal)
                .style(paddingType, this.getLineIndent(paddingVal));
            this.finalTextContainer = this.finalTextContainer.append("div").classed("tw_pers", true);
        }
        public textSettingsBottom(textSettings: ITextSettings): void {
            const paddingVal: number = textSettings.lineIndent >= 0 ?
                textSettings.lineIndent : (-textSettings.lineIndent);
            let transformed: string = "";
            let paddingType: string = "";
            let positionName: string = "";
            let positionVal: string = "";
            const propVal: string = this.finalTextContainer.style("transform");
            positionName = "top";
            positionVal = "100%";
            switch (textSettings.direction) {
                case "horizontal-tb": {
                    transformed = `${propVal}  translate(0%, -100%)`;
                    paddingType = textSettings.lineIndent >= 0 ?
                        "padding-top" : "padding-bottom";
                    if (textSettings.alignment === "right") {
                        this.finalTextContainer.style("float", "right");
                    } else if (textSettings.alignment === "center") {
                        this.finalTextContainer.style("width", "");
                    }
                }
                                      break;
                case "horizontal-bt": {
                    transformed = `${propVal}  translate(0%, 100%)`;
                    paddingType = textSettings.lineIndent >= 0 ?
                        "padding-bottom" : "padding-top";

                    if (textSettings.alignment === "right") {
                        this.finalTextContainer.style("float", "right");
                    } else if (textSettings.alignment === "center") {
                        this.finalTextContainer.style("width", "");
                    }
                }
                                      break;
                case "vertical-rl": {
                    if (textSettings.alignment === "center") {
                        transformed = `${propVal}  translate(-50%, -100%)`;
                        paddingType = textSettings.lineIndent >= 0 ?
                            "padding-right" : "padding-left";
                        this.finalTextContainer = this.finalTextContainer
                            .style("left", "50%");
                    } else if (textSettings.alignment === "right") {
                        transformed = `${propVal}  translate(-100%, -100%)`;
                        paddingType = textSettings.lineIndent >= 0 ?
                            "padding-right" : "padding-left";
                        this.finalTextContainer = this.finalTextContainer
                            .style("left", "100%");
                    } else if (textSettings.alignment === "left") {
                        transformed = `${propVal}  translate(0%, -100%)`;
                        paddingType = textSettings.lineIndent >= 0 ?
                            "padding-right" : "padding-left";
                    }
                }
                                    break;
                case "vertical-lr": {
                    if (textSettings.alignment === "center") {
                        transformed = `${propVal}  translate(50%, 100%)`;
                        paddingType = textSettings.lineIndent >= 0 ?
                            "padding-left" : "padding-right";
                        this.finalTextContainer = this.finalTextContainer
                            .style("left", "50%");
                    } else if (textSettings.alignment === "right") {
                        transformed = `${propVal}  translate(100%, 100%)`;
                        paddingType = textSettings.lineIndent >= 0 ?
                            "padding-left" : "padding-right";
                        this.finalTextContainer = this.finalTextContainer
                            .style("left", "100%");
                    } else if (textSettings.alignment === "left") {
                        transformed = `${propVal}  translate(0%, 100%)`;
                        paddingType = textSettings.lineIndent >= 0 ?
                            "padding-left" : "padding-right";
                    }
                }
                                    break;
                default: break;
            }
            this.finalTextContainer = this.finalTextContainer
                .style("position", "relative")
                .style("transform", transformed)
                .style(positionName, positionVal)
                .style(paddingType, this.getLineIndent(paddingVal));
            this.finalTextContainer = this.finalTextContainer.append("div").classed("tw_pers", true);
        }

        public textSettings(textSettings: ITextSettings): void {
            switch (textSettings.alignmentV) {
                case "top": {
                    this.textSettingsTop(textSettings);
                }           break;

                case "middle": {
                    this.textSettingsMiddle(textSettings);
                }
                               break;

                case "bottom": {
                    this.textSettingsBottom(textSettings);
                }              break;
                default: break;
            }
        }
        public handleHeightIssue(options: VisualUpdateOptions,
                                 dynamicText: any, url: any, valueLength: number, textSettings: ITextSettings): void {
            const upper = 100;
            const pers: number = textSettings.perspective > 0 ? upper - textSettings.perspective + 1 : 0;
            if (textSettings.direction === "vertical-lr" || textSettings.direction === "vertical-rl") {
                this.finalTextContainer.style("max-height", `${$("#sandbox-host").height()}px`);
            } else if (textSettings.direction === "horizontal-tb" || textSettings.direction === "horizontal-bt") {
                this.finalTextContainer.style("max-width", `${$("#sandbox-host").width()}px`);
            }
            if (valueLength === 1) {
                let transformedVal: string = "";
                if (pers === null || pers === 0) {
                    d3.select(".tw_finalText").style("perspective", "none");
                } else {
                    d3.select(".tw_finalText")
                        .style("perspective", this.getPerspective(pers))
                        .style("perspective-origin", "center")
                        .attr("overflow-x", "visible");
                    if (textSettings.direction === "vertical-rl" || textSettings.direction === "vertical-lr") {
                        transformedVal = "rotateY(25deg)";
                    } else {
                        transformedVal = "rotateX(25deg)";
                    }
                }
                d3.select(".tw_pers").style("transform", transformedVal);
            }
            // Below Two lines are to handle height issue of div in edge
            const dataView: DataView = this.dataViews = options.dataViews[0];
            const spanHeight: number = dynamicText.height();
            $(".tw_value.tw_finalText").height(spanHeight + 2);
            for (const jIterator of options.dataViews[0].categorical.categories) {
                if (jIterator.source.type[`category`] === "WebUrl" && jIterator.source.roles.URL) {
                    dynamicText.on("click", (): void => {
                        this.visualHost.launchUrl(url);
                    });
                }
            }
            if (dataView.categorical.categories !== undefined) {
                for (const iterator of options.dataViews[0].categorical.categories) {
                    if (iterator.source.type[`category`]
                        === "WebUrl" && iterator.source.roles.URL) {
                        dynamicText.addClass("urlIcon");
                    }
                }
            } else {
                if (dataView.categorical.values[0].source.roles.URL) {
                    if (dataView.categorical.values[0].source.type[`category`] === "WebUrl") {
                        dynamicText.addClass("urlIcon");
                    }
                }
            }
            this.eventService.renderingFinished(options);
        }
        public textOrdering(textValStatic: string, textValDynamic: string,
                            textFontSize: number, dataView: DataView): void {
            const dynFontStyleClass: string = this.getFontStyleClassName(this.dynamicSettings);
            const dynTextDecoration: string = this.getTextDecoration(this.dynamicSettings);
            const dynTextShadow: string = this.getDynamicTextSettings(dataView).textShadow;
            const dynTextShadowBlur: string = this.getDynamicTextSettings(dataView).textShadowBlur;
            const dynTextShadowColor: string = this.getDynamicTextSettings(dataView).textShadowColor;
            const staticFontStyleClass: string = this.getFontStyleClassName(this.staticTextSettings);
            const staticTextDecoration: string = this.getTextDecoration(this.staticTextSettings);
            const staticTextShadow: string = this.getStaticTextSettings(dataView).textShadow;
            const staticTextShadowBlur: string = this.getStaticTextSettings(dataView).textShadowBlur;
            const staticTextShadowColor: string = this.getStaticTextSettings(dataView).textShadowColor;
            const staticFontWgt: string = this.getFontWeight(this.staticTextSettings);
            const textTrans: string = this.getTextTransform(this.staticTextSettings);
            const staticTextFontFamily: string = this.staticTextSettings.fontFamily;
            const textTransD: string = this.getTextTransform(this.dynamicSettings);
            const dynFontWgt: string = this.getFontWeight(this.dynamicSettings);
            const dynamicTextFontFamily: string = this.dynamicSettings.fontFamily;
            const colonText: string = " : ";
            if (textValStatic !== "" && this.staticTextSettings.showColon) {
                if (this.staticTextSettings.textPosition === "suffix") {
                    this.getText(textValDynamic, dynFontStyleClass, dynTextDecoration,
                        textFontSize, dynTextShadow, dynTextShadowBlur,
                        dynTextShadowColor, dynamicTextFontFamily,
                        this.dynamicSettings.backgroundColor, dynFontWgt, textTransD);
                    this.colonText(colonText);
                    this.getTexts(textValStatic, staticFontStyleClass,
                        staticTextDecoration, textFontSize, staticTextFontFamily,
                        this.staticTextSettings.backgroundColor, textTrans,
                        staticTextShadow, staticTextShadowBlur,
                        staticTextShadowColor, staticFontWgt);
                    if (this.dynamicSettings.italicStyle) {
                        $(".dynamicpluscolon").css("padding-left", "4px");
                    }
                } else {
                    this.getTexts(textValStatic, staticFontStyleClass, staticTextDecoration,
                        textFontSize, staticTextFontFamily, this.staticTextSettings.backgroundColor,
                        textTrans, staticTextShadow, staticTextShadowBlur, staticTextShadowColor, staticFontWgt);
                    this.colonText(colonText);
                    this.getText(textValDynamic, dynFontStyleClass, dynTextDecoration, textFontSize,
                        dynTextShadow, dynTextShadowBlur, dynTextShadowColor, dynamicTextFontFamily,
                        this.dynamicSettings.backgroundColor, dynFontWgt, textTransD);
                    if (this.staticTextSettings.italicStyle) {
                        $(".dynamicpluscolon").css("padding-left", "4px");
                    }
                }
            } else if (textValStatic !== "" && !this.staticTextSettings.showColon) {
                if (this.staticTextSettings.textPosition === "suffix") {
                    this.getText(textValDynamic, dynFontStyleClass, dynTextDecoration, textFontSize, dynTextShadow,
                        dynTextShadowBlur, dynTextShadowColor, dynamicTextFontFamily,
                        this.dynamicSettings.backgroundColor, dynFontWgt, textTransD);
                    this.addSpace();
                    this.getTexts(textValStatic, staticFontStyleClass, staticTextDecoration, textFontSize,
                        staticTextFontFamily, this.staticTextSettings.backgroundColor, textTrans, staticTextShadow,
                        staticTextShadowBlur, staticTextShadowColor, staticFontWgt);
                } else {
                    this.getTexts(textValStatic, staticFontStyleClass, staticTextDecoration, textFontSize,
                        staticTextFontFamily, this.staticTextSettings.backgroundColor, textTrans, staticTextShadow,
                        staticTextShadowBlur, staticTextShadowColor, staticFontWgt);
                    this.addSpace();
                    this.getText(textValDynamic, dynFontStyleClass, dynTextDecoration, textFontSize,
                        dynTextShadow, dynTextShadowBlur, dynTextShadowColor, dynamicTextFontFamily,
                        this.dynamicSettings.backgroundColor, dynFontWgt, textTransD);
                }
            } else if (textValStatic === "") {
                this.getText(textValDynamic, dynFontStyleClass, dynTextDecoration, textFontSize, dynTextShadow,
                    dynTextShadowBlur, dynTextShadowColor,
                    dynamicTextFontFamily, this.dynamicSettings.backgroundColor,
                    dynFontWgt, textTransD);
            }
        }
        public update(options: VisualUpdateOptions): void {
            try {
                this.eventService.renderingStarted(options);
                const textDynamic = (".dynamicText");
                this.target.selectAll(".tw_value").remove();
                const dataView: DataView = this.dataViews = options.dataViews[0];
                let valueLength: number = 0;
                const textSettings: ITextSettings = this.getTextSettings(dataView);
                this.dynamicSettings = this.getDynamicTextSettings(dataView);
                this.staticTextSettings = this.getStaticTextSettings(dataView);
                const textValStaticInput: string = this.staticTextSettings.postText;
                this.staticTextSettings.postText = textValStaticInput;
                const valuesContainer: IDynamicTextContainer = this.getDynamicTextValue(dataView);
                const textValDynamicInput: string = valuesContainer.textContainer;
                const textFontSize: number = textSettings.fontSize;
                const letSpacing: number = textSettings.letterSpacing;
                const wordSpace: number = textSettings.wordSpacing;
                const lHeight: number = textSettings.lineHeight;
                const pers: number = textSettings.perspective > 0 ? 100 - textSettings.perspective + 1 : 0;
                const indent: number = textSettings.textIndent;
                let textRotationVal: number = textSettings.textRotate === null ? 0 : textSettings.textRotate;
                const textSkewX: number = textSettings.skewX;
                const textSkewY: number = textSettings.skewY;
                let textValStatic: string = "";
                let textValDynamic: string = "";
                valueLength = valuesContainer.lengthContainer;
                if (valueLength === 1) { // Text Formatting
                    const original: d3.Selection<HTMLElement> = this.target.append("div")
                        .classed("tw_value tw_finalText", true)
                        .style("font-size", this.pointToPixel(textFontSize))
                        .style("letter-spacing", this.letSpace(letSpacing))
                        .style("word-spacing", this.getWordSpace(wordSpace))
                        .style("line-height", this.getLineHeight(lHeight))
                        .style("text-indent", this.getTextIndent(indent))
                        .style("color", textSettings.color +
                            this.getOpacityHex(textSettings.transparency == null ? 0 : textSettings.transparency))
                        .style("transform", this.getSkewString(textSkewX, textSkewY)).style("width", "fit-content");
                    textValStatic = textValStaticInput;
                    textValDynamic = textValDynamicInput;
                } else {
                    let errMsg: string = "";
                    if (valueLength > 1) {
                        errMsg = "Query returned more than one row, please filter data to return one row";
                    } else if (valueLength === 0) { errMsg = "Query contains null value"; }
                    const original: d3.Selection<HTMLElement> = this.target.append("div")
                        .classed("tw_value errormsg", true).text(errMsg).attr("title", errMsg)
                        .style("font-size", this.pointToPixel(textFontSize))
                        .style("letter-spacing", this.letSpace(letSpacing))
                        .style("word-spacing", this.getWordSpace(wordSpace))
                        .style("line-height", this.getLineHeight(lHeight))
                        .style("text-indent", this.getTextIndent(indent))
                        .style("perspective", this.getPerspective(pers))
                        .style("font-family", "Segoe UI Semibold")
                        .style("color", "#777")
                        .style("transform", this.getSkewString(textSkewX, textSkewY));
                }
                let url;
                if (dataView.categorical.categories !== undefined) { // To check if url field exists
                    for (const iterator of options.dataViews[0].categorical.categories) {
                        if (iterator.source.type[`category`] === "WebUrl" && iterator.source.roles.URL) {
                            url = (iterator.values.toString());
                        }
                    }
                } else {
                    if (dataView.categorical.values[0].source.roles.URL) {
                        if (dataView.categorical.values[0].source.type[`category`]
                            === "WebUrl") {
                            url = (dataView.categorical.values[0].values.toString());
                        }
                    }
                }// Text Direction
                let textAlign: string = textSettings.alignment;
                let writingMode: string = textSettings.direction;
                switch (textSettings.direction) {
                    case "vertical-lr":
                        textRotationVal = this.degree180 + textRotationVal; writingMode = "tb-rl";
                        break;
                    case "vertical-rl":
                        writingMode = "tb-rl"; break;
                    case "horizontal-bt":
                        textRotationVal = this.degree180 + textRotationVal;
                        textAlign = textAlign === "left" ? "right" : (textAlign === "right" ? "left" : "center");
                        writingMode = "horizontal-tb"; break;
                    default: break;
                }
                const element: string = d3.select(".tw_finalText").style("transform");
                const newTransform: string = `${element} rotate(${textRotationVal}deg)`;
                this.finalTextContainer = d3.select(".tw_finalText")
                    .style("text-align", textAlign).style("writing-mode", writingMode).style("transform", newTransform);
                this.textSettings(textSettings); // Vertical Alignment & Line Indentation
                this.textOrdering(textValStatic, textValDynamic, textFontSize, dataView); // Text Ordering
                const dynamicText = $(".dynamicText");
                const twFinalText = $(".tw_finalText");
                this.textOverflow(textRotationVal, textSettings, dynamicText, twFinalText); // Text Overflow Handling
                this.handleHeightIssue(options, dynamicText, url, valueLength, textSettings); // Applying Perpective
            } catch (exception) { this.eventService.renderingFailed(options, exception); }
        }
        public textOverflowMiddleRL(textRotationVal: number, textSettings: ITextSettings,
                                    textHeight: number, textWidth: number): void {
            let marginT: number = 0;
            let marginL: number = 0;
            let buffer: number = 0;
            let rotVal: number = textRotationVal > 0 ? textRotationVal : -textRotationVal;
            const upper = 100;
            rotVal = rotVal % this.degree360;
            if (rotVal < this.degree180) {
                buffer = (rotVal / upper * 2) * textWidth;
            } else {
                buffer = ((this.degree360 - rotVal) / upper * 2) * textWidth;
            }
            switch (textSettings.alignment) {
                case "left": {
                    if (rotVal > 0 && rotVal <= this.degree90) {
                        const rotationHeight: number = Math.sin(this.toRadians(this.degree90 - rotVal));
                        marginT = ((textHeight - (textHeight) * rotationHeight) / 2);
                    } else if ((rotVal > this.degree90 && rotVal <= this.degree270)) {
                        const rotationHeight: number = Math.sin(this.toRadians(rotVal - this.degree90));
                        marginT = ((textHeight + (textHeight) * rotationHeight) / 2);
                    } else if (rotVal > this.degree270 && rotVal < this.degree360) {
                        const rotationHeight: number = Math
                            .sin(this.toRadians(rotVal % this.degree270));
                        marginT = ((textHeight - (textHeight) * rotationHeight) / 2);
                    }
                    if ((rotVal > this.degree180 && rotVal < this.degree360)) {
                        const rotationHeight: number = -Math.sin(this.toRadians(rotVal));
                        marginL = (((textHeight) * rotationHeight));
                    }
                    this.finalTextContainer.style("margin-top", `${-marginT}px`);
                    this.finalTextContainer.style("margin-left", `${marginL - buffer}px`);
                }
                             break;
                case "right": {
                    if (rotVal > 0 && rotVal <= this.degree90) {
                        const rotationHeight: number = Math.sin(this.toRadians(this.degree90 - rotVal));
                        marginT = ((textHeight - (textHeight) * rotationHeight) / 2);
                    } else if ((rotVal > this.degree90 && rotVal <= this.degree270)) {
                        const rotationHeight: number = Math.sin(this.toRadians(rotVal - this.degree90));
                        marginT = ((textHeight + (textHeight) * rotationHeight) / 2);
                    } else if (rotVal > this.degree270 && rotVal < this.degree360) {
                        const rotationHeight: number = Math
                            .sin(this.toRadians(rotVal % this.degree270));
                        marginT = ((textHeight - (textHeight) * rotationHeight) / 2);
                    }
                    if ((rotVal > 0 && rotVal < this.degree180)) {
                        const rotationHeight: number = -Math.sin(this.toRadians(rotVal));
                        marginL = (((textHeight) * rotationHeight));
                    }
                    this.finalTextContainer.style("margin-top", `${-marginT}px`);
                    this.finalTextContainer.style("margin-left", `${marginL - buffer}px`);
                }
                              break;
                case "center": {
                    if (rotVal > 0 && rotVal <= this.degree90) {
                        const rotationHeight: number = Math.sin(this.toRadians(this.degree90 - rotVal));
                        marginT = ((textHeight - (textHeight) * rotationHeight) / 2);
                    } else if ((rotVal > this.degree90 && rotVal <= this.degree270)) {
                        const rotationHeight: number = Math.sin(this.toRadians(rotVal - this.degree90));
                        marginT = ((textHeight + (textHeight) * rotationHeight) / 2);
                    } else if (rotVal > this.degree270 && rotVal < this.degree360) {
                        const rotationHeight: number = Math
                            .sin(this.toRadians(rotVal % this.degree270));
                        marginT = ((textHeight - (textHeight) * rotationHeight) / 2);
                    }
                    let rotValHeight: number = Math.sin(this.toRadians(rotVal));
                    rotValHeight = rotValHeight > 0 ? rotValHeight : -rotValHeight;
                    marginL = (((textHeight) * rotValHeight)) / 2;
                    this.finalTextContainer.style("margin-top", `${-marginT}px`);
                    if (rotVal < this.degree180) {
                        this.finalTextContainer.style("margin-left", `${-marginL}px`);
                    } else {
                        this.finalTextContainer.style("margin-left", `${marginL}px`);
                    }
                }
                               break;
                default: break;
            }
        }
        public textOverflowMiddle(textRotationVal: number, textSettings: ITextSettings,
                                  textHeight: number, textWidth: number): void {
            let marginT: number = 0;
            let marginL: number = 0;
            let buffer: number = 0;
            let rotVal: number = textRotationVal > 0 ? textRotationVal : -textRotationVal;
            switch (textSettings.direction) {
                case "vertical-rl": {
                    this.textOverflowMiddleRL(textRotationVal, textSettings, textHeight, textWidth);
                }                   break;

                case "vertical-lr": {
                    rotVal = rotVal - this.degree180;
                    rotVal = rotVal % this.degree360;
                    const upper = 100;
                    if (rotVal < this.degree180) {
                        buffer = (rotVal / upper * 2) * textSettings.fontSize;
                    } else {
                        buffer = ((this.degree360 - rotVal) / upper * 2) * textSettings.fontSize;
                    }
                    switch (textSettings.alignment) {
                        case "left": {
                            if (rotVal > 0 && rotVal <= this.degree90) {
                                const rotValHei: number = Math
                                    .sin(this.toRadians(this.degree90 - rotVal % this.degree90));
                                marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                            } else if ((rotVal > this.degree90 && rotVal <= this.degree270)) {
                                const rotValHei: number = Math.sin(this.toRadians(rotVal - this.degree90));
                                marginT = ((textHeight + (textHeight) * rotValHei) / 2);
                            } else if (rotVal > this.degree270 && rotVal < this.degree360) {
                                const rotValHei: number = Math.sin(this.toRadians(rotVal % this.degree270));
                                marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                            }
                            if ((rotVal > this.degree180 && rotVal < this.degree360)) {
                                const rotValHei: number = -Math.sin(this.toRadians(rotVal));
                                marginL = (((textHeight) * rotValHei));
                            }
                            this.finalTextContainer.style("margin-top", `${-marginT}px`);
                            this.finalTextContainer.style("margin-left", `${marginL}px`);
                        }
                                     break;
                        case "right": {
                            if (rotVal > 0 && rotVal <= this.degree90) {
                                const rotValHei: number = Math.sin(this.toRadians(this.degree90 - rotVal));
                                marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                            } else if ((rotVal > this.degree90 && rotVal <= this.degree270)) {
                                const rotValHei: number = Math.sin(this.toRadians(rotVal - this.degree90));
                                marginT = ((textHeight + (textHeight) * rotValHei) / 2);
                            } else if (rotVal > this.degree270 && rotVal < this.degree360) {
                                const rotValHei: number = Math.sin(this.toRadians(rotVal % this.degree270));
                                marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                            }
                            if ((rotVal > 0 && rotVal < this.degree180)) {
                                const rotValHei: number = -Math.sin(this.toRadians(rotVal));
                                marginL = (((textHeight) * rotValHei));
                            }
                            this.finalTextContainer.style("margin-top", `${-marginT}px`);
                            this.finalTextContainer.style("margin-left", `${(marginL - buffer)}px`);
                        }
                                      break;
                        case "center": {
                            let rotValHei: number;
                            if (rotVal > 0 && rotVal <= this.degree90) {
                                rotValHei = Math
                                    .sin(this.toRadians(this.degree90 - rotVal % this.degree90));
                                marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                            } else if ((rotVal > this.degree90 && rotVal <= this.degree270)) {
                                rotValHei = Math.sin(this.toRadians(rotVal - this.degree90));
                                marginT = ((textHeight + (textHeight) * rotValHei) / 2);
                            } else if (rotVal > this.degree270 && rotVal < this.degree360) {
                                rotValHei = Math.sin(this.toRadians(rotVal % this.degree270));
                                marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                            }
                            rotValHei = Math.sin(this.toRadians(rotVal));
                            rotValHei = rotValHei > 0 ? rotValHei : -rotValHei;
                            marginL = (((textHeight) * rotValHei)) / 2;
                            this.finalTextContainer.style("margin-top", `${-marginT}px`);
                            if (rotVal < this.degree180) {
                                this.finalTextContainer
                                    .style("margin-left", `${-marginL}px`);
                            } else {
                                this.finalTextContainer.style("margin-left", `${marginL}px`);
                            }
                        }
                                       break;
                        default: break;
                    }
                }                   break;
                default: break;
            }
        }
        public textOverflowBottomRL(textRotationVal: number, textSettings: ITextSettings, textHeight: number): void {
            let marginT: number = 0;
            let marginL: number = 0;
            let buffer: number = 0;
            let rotVal: number = textRotationVal > 0 ? textRotationVal : -textRotationVal;
            rotVal = rotVal % this.degree360;
            const upper = 100;
            if (rotVal < this.degree180) {
                buffer = (rotVal / upper * 2) * textSettings.fontSize;
            } else {
                buffer = ((this.degree360 - rotVal) / upper * 2) * textSettings.fontSize;
            }
            switch (textSettings.alignment) {
                case "left": {
                    if (rotVal > 0 && rotVal <= this.degree90) {
                        const rotValHei: number = Math.sin(this.toRadians(this.degree90 - rotVal));
                        marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                    } else if ((rotVal > this.degree90 && rotVal <= this.degree270)) {
                        const rotValHei: number = Math.sin(this.toRadians(rotVal - this.degree90));
                        marginT = ((textHeight + (textHeight) * rotValHei) / 2);

                    } else if (rotVal > this.degree270 && rotVal < this.degree360) {
                        const rotValHei: number = Math.sin(this.toRadians(rotVal % this.degree270));
                        marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                    }
                    if (rotVal <= this.degree180) {
                        const rotValHei: number = -Math.sin(this.toRadians(rotVal % this.degree180));
                        marginL = ((textHeight) * rotValHei) / 2;
                    } else if (rotVal > this.degree180) {
                        const rotValHei: number = Math.sin(this.toRadians(rotVal % this.degree180));
                        marginL = ((1.5 * textHeight) * rotValHei);
                    }
                    this.finalTextContainer.style("margin-top", `${-2 * marginT}px`);
                    this.finalTextContainer.style("margin-left", `${marginL}px`);
                }
                             break;
                case "right": {
                    if (rotVal > 0 && rotVal <= this.degree90) {
                        const rotValHei: number = Math.sin(this.toRadians(this.degree90 - rotVal));
                        marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                    } else if ((rotVal > this.degree90 && rotVal <= this.degree270)) {
                        const rotValHei: number = Math.sin(this.toRadians(rotVal - this.degree90));
                        marginT = ((textHeight + (textHeight) * rotValHei) / 2);
                    } else if (rotVal > this.degree270 && rotVal < this.degree360) {
                        const rotValHei: number = Math.sin(this.toRadians(rotVal % this.degree270));
                        marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                    }
                    if (rotVal <= this.degree180) {
                        const rotValHei: number = -Math.sin(this.toRadians(rotVal % this.degree180));
                        marginL = ((1.5 * textHeight) * rotValHei);

                    } else if (rotVal > this.degree180) {
                        const rotValHei: number = Math.sin(this.toRadians(rotVal % this.degree180));
                        marginL = ((textHeight) * rotValHei) / 2;
                    }
                    this.finalTextContainer.style("margin-top", `${-2 * marginT}px`);
                    this.finalTextContainer
                        .style("margin-left", `${marginL - buffer}px`);
                }
                              break;
                case "center": {
                    let rotValHei: number;
                    if (rotVal > 0 && rotVal <= this.degree90) {
                        rotValHei = Math.sin(this.toRadians(this.degree90 - rotVal));
                        marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                    } else if ((rotVal > this.degree90 && rotVal <= this.degree270)) {
                        rotValHei = Math.sin(this.toRadians(rotVal - this.degree90));
                        marginT = ((textHeight + (textHeight) * rotValHei) / 2);
                    } else if (rotVal > this.degree270 && rotVal < this.degree360) {
                        rotValHei = Math.sin(this.toRadians(rotVal % this.degree270));
                        marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                    }
                    if (rotVal <= this.degree180) {
                        rotValHei = -Math.sin(this.toRadians(rotVal % this.degree180));
                        marginL = ((textHeight) * rotValHei) / 2;
                    } else if (rotVal > this.degree180) {
                        rotValHei = Math.sin(this.toRadians(rotVal % this.degree180));
                        marginL = ((1.5 * textHeight) * rotValHei);
                    }
                    rotValHei = Math.sin(this.toRadians(rotVal));
                    rotValHei = rotValHei > 0 ? rotValHei : -rotValHei;
                    marginL = (((textHeight) * rotValHei));
                    this.finalTextContainer
                        .style("margin-top", `${-2 * marginT}px`);
                    if (rotVal < this.degree180) {
                        this.finalTextContainer
                            .style("margin-left", `${-marginL}px`);
                    } else {
                        this.finalTextContainer
                            .style("margin-left", `${marginL}px`);
                    }
                }
                               break;
                default: break;
            }
        }
        public textOverflowBottomLR(textRotationVal: number, textSettings: ITextSettings,
                                    textHeight: number, textWidth: number, textWidth2: number): void {
            let marginT: number = 0;
            let marginL: number = 0;
            let buffer: number = 0;
            let rotVal: number = textRotationVal > 0 ? textRotationVal : -textRotationVal;
            rotVal = rotVal - this.degree180;
            rotVal = rotVal % this.degree360;
            const upper = 100;
            if (rotVal < this.degree180) {
                buffer = (rotVal / upper * 2) * textSettings.fontSize;
            } else {
                buffer = ((this.degree360 - rotVal) / upper * 2) * textSettings.fontSize;
            }
            switch (textSettings.alignment) {
                case "left": {
                    if (rotVal > 0 && rotVal <= this.degree90) {
                        const rotValHei: number = Math.sin(this.toRadians(this.degree90 - rotVal));
                        marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                    } else if ((rotVal > this.degree90 && rotVal <= this.degree270)) {
                        const rotValHei: number = Math.sin(this.toRadians(rotVal - this.degree90));
                        marginT = ((textHeight + (textHeight) * rotValHei) / 2);
                    } else if (rotVal > this.degree270 && rotVal < this.degree360) {
                        const rotValHei: number = Math.sin(this.toRadians(rotVal % this.degree270));
                        marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                    }
                    if (rotVal <= this.degree180) {
                        const rotValHei: number = -Math.sin(this.toRadians(rotVal % this.degree180));
                        marginL = ((textHeight) * rotValHei) / 2;
                    } else if (rotVal > this.degree180) {
                        const rotValHei: number = Math.sin(this.toRadians(rotVal % this.degree180));
                        marginL = ((1.5 * textHeight) * rotValHei);
                    }
                    this.finalTextContainer.style("margin-top", `${-2 * marginT}px`);
                    this.finalTextContainer.style("margin-left", `${marginL}px`);
                }
                             break;
                case "right": {
                    if (rotVal > 0 && rotVal <= this.degree90) {
                        const rotValHei: number = Math.sin(this.toRadians(this.degree90 - rotVal));
                        marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                    } else if ((rotVal > this.degree90 && rotVal <= this.degree270)) {
                        const rotValHei: number = Math.sin(this.toRadians(rotVal - this.degree90));
                        marginT = ((textHeight + (textHeight) * rotValHei) / 2);

                    } else if (rotVal > this.degree270 && rotVal < this.degree360) {
                        const rotValHei: number = Math.sin(this.toRadians(rotVal % this.degree270));
                        marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                    }
                    if (rotVal <= this.degree180) {
                        const rotValHei: number = -Math.sin(this.toRadians(rotVal % this.degree180));
                        marginL = ((1.5 * textHeight) * rotValHei);

                    } else if (rotVal > this.degree180) {
                        const rotValHei: number = Math.sin(this.toRadians(rotVal % this.degree180));
                        marginL = ((textHeight) * rotValHei) / 2;
                    }
                    this.finalTextContainer.style("margin-top", `${-2 * marginT}px`);
                    this.finalTextContainer
                        .style("margin-left", `${marginL - buffer}px`);
                }
                              break;
                case "center": {
                    let rotValHei: number;
                    if (rotVal > 0 && rotVal <= this.degree90) {
                        rotValHei = Math.sin(this.toRadians(this.degree90 - rotVal));
                        marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                    } else if ((rotVal > this.degree90 && rotVal <= this.degree270)) {
                        rotValHei = Math.sin(this.toRadians(rotVal - this.degree90));
                        marginT = ((textHeight + (textHeight) * rotValHei) / 2);

                    } else if (rotVal > this.degree270 && rotVal < this.degree360) {
                        rotValHei = Math.sin(this.toRadians(rotVal % this.degree270));
                        marginT = ((textHeight - (textHeight) * rotValHei) / 2);
                    }
                    if (rotVal <= this.degree180) {
                        rotValHei = -Math.sin(this.toRadians(rotVal % this.degree180));
                        marginL = ((textHeight) * rotValHei) / 2;
                    } else if (rotVal > this.degree180) {
                        rotValHei = Math.sin(this.toRadians(rotVal % this.degree180));
                        marginL = ((1.5 * textHeight) * rotValHei);
                    }
                    rotValHei = Math.sin(this.toRadians(rotVal));
                    rotValHei = rotValHei > 0 ? rotValHei : -rotValHei;
                    marginL = (((textHeight) * rotValHei));
                    this.finalTextContainer
                        .style("margin-top", `${-2 * marginT}px`);
                    if (rotVal < this.degree180) {
                        this.finalTextContainer
                            .style("margin-left", `${-marginL}px`);
                    } else {
                        this.finalTextContainer
                            .style("margin-left", `${marginL}px`);
                    }
                }
                               break;
                default: break;
            }
        }
        public textOverflowBottom(textRotationVal: number, textSettings: ITextSettings,
                                  textHeight: number, textWidth: number, textWidth2: number): void {
            const upper = 100;
            switch (textSettings.direction) {
                case "horizontal-tb": {
                    let buffer: number = 0;
                    let rotVal: number = textRotationVal > 0 ?
                        textRotationVal : -textRotationVal;
                    textRotationVal = textRotationVal > 0 ? textRotationVal % this.degree180 :
                        (-textRotationVal) % this.degree180;
                    rotVal = rotVal % this.degree360;
                    if (rotVal < this.degree180) {
                        buffer = (rotVal / upper * 2) * textSettings.fontSize;
                    } else {
                        buffer =
                            ((this.degree360 - rotVal) / upper * 2) * textSettings.fontSize;
                    }
                    if (textSettings.alignment !== "center") {
                        this.finalTextContainer
                            .style("margin-top", `${(-((textWidth / 2) * Math.sin(
                                this.toRadians(textRotationVal)) + (buffer)))}px`);
                    } else {
                        this.finalTextContainer
                            .style("margin-top", `${(-((textWidth2 / 2) * Math.sin(
                                this.toRadians(textRotationVal)) + (buffer)))}px`);
                    }
                }
                                      break;
                case "horizontal-bt": {
                    let buffer: number = 0;
                    textRotationVal = textRotationVal - this.degree180;
                    const rotVal: number = textRotationVal > 0 ?
                        textRotationVal : -textRotationVal;
                    textRotationVal = textRotationVal > 0 ? textRotationVal % this.degree180 :
                        (-textRotationVal) % this.degree180;
                    if (rotVal < this.degree180) {
                        buffer = (rotVal / upper * 2) * textSettings.fontSize;
                    } else {
                        buffer = ((this.degree360 - rotVal) / upper * 2) * textSettings.fontSize;
                    }

                    if (textSettings.alignment !== "center") {
                        this.finalTextContainer
                            .style("margin-top", `${(-((textWidth / 2) *
                                Math.sin(this.toRadians(textRotationVal)) + (buffer)))}px`);
                    } else {
                        this.finalTextContainer
                            .style("margin-top", `${(-((textWidth2 / 2) *
                                Math.sin(this.toRadians(textRotationVal)) + (buffer)))}px`);
                    }
                }
                                      break;
                case "vertical-rl": {
                    this.textOverflowBottomRL(textRotationVal, textSettings, textHeight);
                }                   break;

                case "vertical-lr": {
                    this.textOverflowBottomLR(textRotationVal, textSettings, textHeight, textWidth, textWidth2);
                }                   break;
                default: break;
            }
        }
        public textOverflow(textRotationVal: number, textSettings: ITextSettings, dynamicText, twFinalText): void {
            const upper = 100;
            if (textRotationVal !== 0) {
                const textWidth: number = twFinalText.width();
                const textWidth2: number = $(".staticText").width() + dynamicText.width() +
                    $(".dynamicpluscolon").width();
                const textHeight: number = twFinalText.height();
                switch (textSettings.alignmentV) {
                    case "top": {
                        switch (textSettings.direction) {
                            case "horizontal-tb": {
                                textRotationVal = textRotationVal > 0 ?
                                    textRotationVal % this.degree180 :
                                    (-textRotationVal) % this.degree180;
                                if (textSettings.alignment !== "center") {
                                    this.finalTextContainer
                                        .style("margin-top", `${((textWidth / 2)
                                            * Math.sin(this.toRadians(textRotationVal)))}px`);
                                } else {
                                    this.finalTextContainer
                                        .style("margin-top", `${((textWidth2 / 2)
                                            * Math.sin(this.toRadians(textRotationVal)))}px`);
                                }
                            }
                                                  break;
                            case "horizontal-bt": {
                                textRotationVal = textRotationVal > 0 ?
                                    textRotationVal % this.degree180 :
                                    (-textRotationVal) % this.degree180;
                                if (textSettings.alignment !== "center") {
                                    this.finalTextContainer.style("margin-top",
                                        `${((textWidth / 2) *
                                            Math.sin(this.toRadians(textRotationVal)))}px`);
                                } else {
                                    this.finalTextContainer
                                        .style("margin-top", `${((textWidth2 / 2)
                                            * Math.sin(this.toRadians(textRotationVal)))}px`);
                                }
                            }
                                                  break;
                            case "vertical-rl": {
                                let buffer: number = 0;
                                let rotVal: number
                                    = textRotationVal > 0 ? textRotationVal : -textRotationVal;
                                textRotationVal = textRotationVal > 0 ?
                                    textRotationVal % this.degree180 :
                                    (-textRotationVal) % this.degree180;
                                rotVal = rotVal % this.degree360;
                                if (rotVal < this.degree180) {
                                    buffer = (rotVal / upper * 2) * textSettings.fontSize;
                                } else {
                                    buffer = ((this.degree360 - rotVal) / upper * 2)
                                        * textSettings.fontSize;
                                }
                                if (textSettings.alignment === "left") {
                                    this.finalTextContainer
                                        .style("margin-left", `${((textHeight / 2)
                                            * Math.sin(this.toRadians(textRotationVal)))}px`);
                                } else if (textSettings.alignment === "right") {
                                    this.finalTextContainer
                                        .style("margin-left", `${-((textHeight / 2) *
                                            Math.sin(this.toRadians(textRotationVal)) + buffer)}px`);
                                }
                            }
                                                break;
                            case "vertical-lr": {
                                let buffer: number = 0;
                                textRotationVal = textSettings.alignment === "right" ?
                                    textRotationVal - this.degree180 : textRotationVal;
                                let rotVal: number =
                                    textRotationVal > 0 ? textRotationVal : -textRotationVal;
                                textRotationVal = textRotationVal > 0 ?
                                    textRotationVal % this.degree180 :
                                    (-textRotationVal) % this.degree180;
                                rotVal = rotVal % this.degree360;
                                if (rotVal < this.degree180) {
                                    buffer = (rotVal / upper * 2) * textSettings.fontSize;
                                } else {
                                    buffer = ((this.degree360 - rotVal) / upper * 2)
                                        * textSettings.fontSize;
                                }
                                if (textSettings.alignment === "left") {
                                    this.finalTextContainer
                                        .style("margin-left", `${((textHeight / 2) *
                                            Math.sin(this.toRadians(textRotationVal)))}px`);
                                } else if (textSettings.alignment === "right") {
                                    this.finalTextContainer
                                        .style("margin-left", `${-((textHeight / 2) *
                                            Math.sin(this.toRadians(textRotationVal)) + buffer)}px`);
                                }
                            }
                                                break;
                            default: break;
                        }
                    }
                                break;
                    case "middle": {
                        this.textOverflowMiddle(textRotationVal, textSettings, textHeight, textWidth);
                    }
                                   break;
                    case "bottom": {
                        this.textOverflowBottom(textRotationVal, textSettings, textHeight, textWidth, textWidth2);
                    }              break;
                    default: break;
                }
            }
        }
        public getDefaultTextSettings(): ITextSettings {
            return {
                alignment: "left",
                alignmentV: "top",
                color: "#000",
                direction: "horizontal-tb",
                fontSize: 18,
                letterSpacing: null,
                lineHeight: null,
                lineIndent: null,
                perspective: null,
                skewX: null,
                skewY: null,
                textIndent: null,
                textRotate: null,
                transparency: null,
                wordSpacing: null
            };
        }

        public getTextSettings(dataView: DataView): ITextSettings {
            let objects: DataViewObjects = null;
            const textSetting: ITextSettings = this.getDefaultTextSettings();
            if (!dataView || !dataView.metadata || !dataView.metadata.objects) {
                return textSetting;
            }
            const upper = 100;
            const lower = -3;
            objects = dataView.metadata.objects;
            textSetting.color = DataViewObjects.getFillColor(objects,
                questTextProperties.textSettings.color, textSetting.color);
            textSetting.transparency = DataViewObjects.getValue(objects, questTextProperties.textSettings.transparency,
                textSetting.transparency) === null ?
                null : (DataViewObjects.getValue(objects,
                    questTextProperties.textSettings.transparency, textSetting.transparency) > upper ?
                    100 : (DataViewObjects.getValue(objects,
                        questTextProperties.textSettings.transparency, textSetting.transparency) < 0 ?
                        0 : DataViewObjects.getValue(objects,
                            questTextProperties.textSettings.transparency, textSetting.transparency)
                    ));
            textSetting.fontSize = DataViewObjects.getValue(objects,
                questTextProperties.textSettings.fontSize, textSetting.fontSize);
            textSetting.alignment = DataViewObjects.getValue(objects,
                questTextProperties.textSettings.alignment, textSetting.alignment);
            textSetting.alignmentV = DataViewObjects.getValue(objects,
                questTextProperties.textSettings.alignmentV, textSetting.alignmentV);
            textSetting.direction = DataViewObjects.getValue(objects,
                questTextProperties.textSettings.direction, textSetting.direction);
            const letSpace = DataViewObjects.getValue(objects,
                questTextProperties.textSettings.letterSpacing, textSetting.letterSpacing) < lower ?
                lower : DataViewObjects.getValue(objects,
                    questTextProperties.textSettings.letterSpacing, textSetting.letterSpacing) > upper / 2 ?
                    upper / 2 : DataViewObjects.getValue(objects,
                        questTextProperties.textSettings.letterSpacing, textSetting.letterSpacing);
            textSetting.letterSpacing = letSpace;
            const wordSpace = DataViewObjects.getValue(objects, questTextProperties.textSettings.wordSpacing,
                textSetting.wordSpacing) < lower ?
                lower : DataViewObjects.getValue(objects,
                    questTextProperties.textSettings.wordSpacing, textSetting.wordSpacing) > upper / 2 ?
                    upper / 2 : DataViewObjects.getValue(objects, questTextProperties.textSettings.wordSpacing,
                        textSetting.wordSpacing);
            textSetting.wordSpacing = wordSpace;
            const lineHeight = DataViewObjects.getValue(objects,
                questTextProperties.textSettings.lineHeight, textSetting.lineHeight) < 0 ?
                0 : DataViewObjects.getValue(objects,
                    questTextProperties.textSettings.lineHeight, textSetting.lineHeight) > upper / 2 ?
                    upper / 2 : DataViewObjects.getValue(objects,
                        questTextProperties.textSettings.lineHeight, textSetting.lineHeight);
            textSetting.lineHeight = lineHeight;
            textSetting.perspective = DataViewObjects.getValue(objects, questTextProperties.textSettings.perspective,
                textSetting.perspective);
            textSetting.perspective = textSetting.perspective === null ?
                null : (textSetting.perspective < 0 ? 0 : textSetting.perspective);
            const textIndent = DataViewObjects.getValue(objects,
                questTextProperties.textSettings.textIndent, textSetting.textIndent) < -3 ?
                -3 : DataViewObjects.getValue(objects,
                    questTextProperties.textSettings.textIndent, textSetting.textIndent);
            textSetting.textIndent = textIndent;
            const lineIndent = DataViewObjects.getValue(objects,
                questTextProperties.textSettings.lineIndent, textSetting.lineIndent) < 0 ?
                0 : DataViewObjects.getValue(objects,
                    questTextProperties.textSettings.lineIndent, textSetting.lineIndent);
            textSetting.lineIndent = lineIndent;
            const getSkew = DataViewObjects.getValue(objects,
                questTextProperties.textSettings.textRotate, textSetting.textRotate) > this.degree360 ?
                this.degree360 : DataViewObjects.getValue(objects,
                    questTextProperties.textSettings.textRotate, textSetting.textRotate) < 0 ?
                    0 : DataViewObjects.getValue(objects,
                        questTextProperties.textSettings.textRotate, textSetting.textRotate);
            textSetting.textRotate = getSkew;
            textSetting.skewX = DataViewObjects.getValue(objects,
                questTextProperties.textSettings.skewX, textSetting.skewX) > this.degree360 ?
                this.degree360 : DataViewObjects.getValue(objects,
                    questTextProperties.textSettings.skewX, textSetting.skewX) < 0 ?
                    0 : DataViewObjects.getValue(objects,
                        questTextProperties.textSettings.skewX, textSetting.skewX);
            textSetting.skewY = DataViewObjects.getValue(objects,
                questTextProperties.textSettings.skewY, textSetting.skewY) > this.degree360 ?
                this.degree360 : DataViewObjects.getValue(objects,
                    questTextProperties.textSettings.skewY, textSetting.skewY) < 0 ?
                    0 : DataViewObjects.getValue(objects,
                        questTextProperties.textSettings.skewY, textSetting.skewY);
            return textSetting;
        }

        public getDefaultStaticTextSettings(): IStaticTextSettings {
            return {
                backgroundColor: "#fff",
                boldStyle: false,
                fontFamily: "Segoe UI",
                fontWeight: "normal",
                italicStyle: false,
                overline: false,
                postText: "",
                showColon: true,
                strikeThrough: false,
                textDecoration: "none",
                textPosition: "prefix",
                textShadow: "none",
                textShadowBlur: "low",
                textShadowColor: "#000",
                textTransform: "",
                transparency: null,
                underline: false
            };
        }

        public getDefaultDynamicTextSettings(): IDynamicTextSettings {
            return {
                backgroundColor: "#FFF",
                boldStyle: false,
                fontFamily: "Segoe UI",
                fontWeight: "normal",
                italicStyle: false,
                overline: false,
                strikeThrough: false,
                textDecoration: "none",
                textShadow: "none",
                textShadowBlur: "low",
                textShadowColor: "#000",
                textTransform: "",
                transparency: null,
                underline: false
            };
        }

        public getDynamicTextSettings(dataView: DataView): IDynamicTextSettings {
            let objects: DataViewObjects = null;
            const dynamicSettings: IDynamicTextSettings = this.getDefaultDynamicTextSettings();
            if (!dataView || !dataView.metadata || !dataView.metadata.objects) {
                return dynamicSettings;
            }
            const upper = 100;
            objects = dataView.metadata.objects;
            dynamicSettings.backgroundColor = DataViewObjects.getFillColor(
                objects, questTextProperties.dynamicSettings.backgroundColor, dynamicSettings.backgroundColor);
            dynamicSettings.transparency = DataViewObjects.getValue(
                objects, questTextProperties.dynamicSettings.transparency,
                dynamicSettings.transparency === null ? 0 : (DataViewObjects.getValue(
                    objects, questTextProperties.dynamicSettings.transparency, dynamicSettings.transparency) > upper ?
                    upper : (DataViewObjects.getValue(
                        objects, questTextProperties.dynamicSettings.transparency, dynamicSettings.transparency) < 0 ?
                        0 : DataViewObjects.getValue(
                            objects, questTextProperties.dynamicSettings.transparency, dynamicSettings.transparency)
                    )));
            dynamicSettings.textDecoration = DataViewObjects.getValue(
                objects, questTextProperties.dynamicSettings.textDecoration, dynamicSettings.textDecoration);
            dynamicSettings.textTransform = DataViewObjects.getValue(
                objects, questTextProperties.dynamicSettings.textTransform, dynamicSettings.textTransform);
            dynamicSettings.textShadow = DataViewObjects.getValue(
                objects, questTextProperties.dynamicSettings.textShadow, dynamicSettings.textShadow);
            dynamicSettings.textShadowBlur = DataViewObjects.getValue(
                objects, questTextProperties.dynamicSettings.textShadowBlur, dynamicSettings.textShadowBlur);
            dynamicSettings.textShadowColor = DataViewObjects.getFillColor(
                objects, questTextProperties.dynamicSettings.textShadowColor, dynamicSettings.textShadowColor);
            dynamicSettings.fontWeight = DataViewObjects.getValue(
                objects, questTextProperties.dynamicSettings.fontWeight, dynamicSettings.fontWeight);
            dynamicSettings.fontFamily = DataViewObjects.getValue(
                objects, questTextProperties.dynamicSettings.fontFamily, dynamicSettings.fontFamily);
            dynamicSettings.boldStyle = DataViewObjects.getValue(
                objects, questTextProperties.dynamicSettings.boldStyle, dynamicSettings.boldStyle);
            dynamicSettings.italicStyle = DataViewObjects.getValue(
                objects, questTextProperties.dynamicSettings.italicStyle, dynamicSettings.italicStyle);
            dynamicSettings.underline = DataViewObjects.getValue(
                objects, questTextProperties.dynamicSettings.underline, dynamicSettings.underline);
            dynamicSettings.overline = DataViewObjects.getValue(
                objects, questTextProperties.dynamicSettings.overline, dynamicSettings.overline);
            dynamicSettings.strikeThrough = DataViewObjects.getValue(
                objects, questTextProperties.dynamicSettings.strikeThrough, dynamicSettings.strikeThrough);

            return dynamicSettings;
        }

        public getStaticTextSettings(dataView: DataView): IStaticTextSettings {
            let objects: DataViewObjects = null;
            const textSetting: IStaticTextSettings = this.getDefaultStaticTextSettings();
            if (!dataView || !dataView.metadata || !dataView.metadata.objects) {
                return textSetting;
            }
            const upper = 100;
            objects = dataView.metadata.objects;
            textSetting.showColon = DataViewObjects.getValue(
                objects, questTextProperties.staticTextSettings.showColon, textSetting.showColon);
            textSetting.textPosition = DataViewObjects.getValue(
                objects, questTextProperties.staticTextSettings.textPosition, textSetting.textPosition);
            textSetting.textDecoration = DataViewObjects.getValue(
                objects, questTextProperties.staticTextSettings.textDecoration, textSetting.textDecoration);
            textSetting.textTransform = DataViewObjects.getValue(
                objects, questTextProperties.staticTextSettings.textTransform, textSetting.textTransform);
            textSetting.textShadow = DataViewObjects.getValue(
                objects, questTextProperties.staticTextSettings.textShadow, textSetting.textShadow);
            textSetting.textShadowBlur = DataViewObjects.getValue(
                objects, questTextProperties.staticTextSettings.textShadowBlur, textSetting.textShadowBlur);
            textSetting.textShadowColor = DataViewObjects.getFillColor(
                objects, questTextProperties.staticTextSettings.textShadowColor, textSetting.textShadowColor);
            textSetting.fontWeight = DataViewObjects.getValue(
                objects, questTextProperties.staticTextSettings.fontWeight, textSetting.fontWeight);
            textSetting.backgroundColor = DataViewObjects.getFillColor(
                objects, questTextProperties.staticTextSettings.backgroundColor, textSetting.backgroundColor);
            textSetting.transparency = DataViewObjects.getValue(
                objects, questTextProperties.staticTextSettings.transparency, textSetting.transparency === null ?
                null : (DataViewObjects.getValue(
                    objects, questTextProperties.staticTextSettings.transparency, textSetting.transparency) > upper ?
                    upper : (DataViewObjects.getValue(
                        objects, questTextProperties.staticTextSettings.transparency, textSetting.transparency) < 0 ?
                        0 : DataViewObjects.getValue(
                            objects, questTextProperties.staticTextSettings.transparency, textSetting.transparency)
                    )
                ));
            textSetting.fontFamily = DataViewObjects.getValue(
                objects, questTextProperties.staticTextSettings.fontFamily, textSetting.fontFamily);
            textSetting.boldStyle = DataViewObjects.getValue(
                objects, questTextProperties.staticTextSettings.boldStyle, textSetting.boldStyle);
            textSetting.italicStyle = DataViewObjects.getValue(
                objects, questTextProperties.staticTextSettings.italicStyle, textSetting.italicStyle);
            textSetting.underline = DataViewObjects.getValue(
                objects, questTextProperties.staticTextSettings.underline, textSetting.underline);
            textSetting.overline = DataViewObjects.getValue(
                objects, questTextProperties.staticTextSettings.overline, textSetting.overline);
            textSetting.strikeThrough = DataViewObjects.getValue(
                objects, questTextProperties.staticTextSettings.strikeThrough, textSetting.strikeThrough);
            textSetting.postText = DataViewObjects.getValue(
                objects, questTextProperties.staticTextSettings.postText, textSetting.postText);

            return textSetting;
        }
        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions):
            VisualObjectInstanceEnumeration {
            const textSetting: ITextSettings = this.getTextSettings(this.dataViews);
            const objectName: string = options.objectName;
            const objectEnumeration: VisualObjectInstance[] = [];
            switch (objectName) {
                case "textSettings":
                    objectEnumeration.push({
                        objectName,
                        properties: {
                            alignment: textSetting.alignment,
                            alignmentV: textSetting.alignmentV,
                            color: textSetting.color,
                            direction: textSetting.direction,
                            fontSize: textSetting.fontSize,
                            letterSpacing: textSetting.letterSpacing,
                            lineHeight: textSetting.lineHeight,
                            lineIndent: textSetting.lineIndent,
                            perspective: textSetting.perspective,
                            skewX: textSetting.skewX,
                            skewY: textSetting.skewY,
                            textIndent: textSetting.textIndent,
                            textRotate: textSetting.textRotate,
                            transparency: textSetting.transparency,
                            wordSpacing: textSetting.wordSpacing
                        },
                        selector: null
                    });
                    break;
                case "staticText":
                    if (this.staticTextSettings.textShadow === "none") {
                        objectEnumeration.push({
                            objectName,
                            properties: {// This field to keep it compatible with the older version. DO NOT DELETE.
                                backgroundColor: this.staticTextSettings.backgroundColor,
                                boldStyle: this.staticTextSettings.boldStyle,
                                fontFamily: this.staticTextSettings.fontFamily,
                                italicStyle: this.staticTextSettings.italicStyle,
                                overline: this.staticTextSettings.overline,
                                postText: this.staticTextSettings.postText,
                                showColon: this.staticTextSettings.showColon,
                                strikeThrough: this.staticTextSettings.strikeThrough,
                                textPosition: this.staticTextSettings.textPosition,
                                textShadow: this.staticTextSettings.textShadow,
                                textTransform: this.staticTextSettings.textTransform,
                                transparency: this.staticTextSettings.transparency,
                                underline: this.staticTextSettings.underline
                            },
                            selector: null
                        });
                    } else {
                        objectEnumeration.push({
                            objectName,
                            properties: {// This field to keep it compatible with the older version. DO NOT DELETE.
                                backgroundColor: this.staticTextSettings.backgroundColor,
                                boldStyle: this.staticTextSettings.boldStyle,
                                fontFamily: this.staticTextSettings.fontFamily,
                                italicStyle: this.staticTextSettings.italicStyle,
                                overline: this.staticTextSettings.overline,
                                postText: this.staticTextSettings.postText,
                                showColon: this.staticTextSettings.showColon,
                                strikeThrough: this.staticTextSettings.strikeThrough,
                                textPosition: this.staticTextSettings.textPosition,
                                textShadow: this.staticTextSettings.textShadow,
                                textShadowBlur: this.staticTextSettings.textShadowBlur,
                                textShadowColor: this.staticTextSettings.textShadowColor,
                                textTransform: this.staticTextSettings.textTransform,
                                transparency: this.staticTextSettings.transparency,
                                underline: this.staticTextSettings.underline,
                            },
                            selector: null
                        });
                    }
                    break;
                case "Settings":
                    if (this.dynamicSettings.textShadow === "none") {
                        objectEnumeration.push({
                            objectName,
                            properties: {
                                backgroundColor: this.dynamicSettings.backgroundColor,
                                boldStyle: this.dynamicSettings.boldStyle,
                                fontFamily: this.dynamicSettings.fontFamily,
                                italicStyle: this.dynamicSettings.italicStyle,
                                overline: this.dynamicSettings.overline,
                                strikeThrough: this.dynamicSettings.strikeThrough,
                                textShadow: this.dynamicSettings.textShadow,
                                textTransform: this.dynamicSettings.textTransform,
                                transparency: this.dynamicSettings.transparency,
                                underline: this.dynamicSettings.underline
                            },
                            selector: null
                        });
                    } else {
                        objectEnumeration.push({
                            objectName,
                            properties: {
                                backgroundColor: this.dynamicSettings.backgroundColor,
                                boldStyle: this.dynamicSettings.boldStyle,
                                fontFamily: this.dynamicSettings.fontFamily,
                                italicStyle: this.dynamicSettings.italicStyle,
                                overline: this.dynamicSettings.overline,
                                strikeThrough: this.dynamicSettings.strikeThrough,
                                textShadow: this.dynamicSettings.textShadow,
                                textShadowBlur: this.dynamicSettings.textShadowBlur,
                                textShadowColor: this.dynamicSettings.textShadowColor,
                                textTransform: this.dynamicSettings.textTransform,
                                transparency: this.dynamicSettings.transparency,
                                underline: this.dynamicSettings.underline
                            },
                            selector: null
                        });
                    }
                    break;
                default: break;
            }

            return objectEnumeration;
        }
        private getTexts(text: string, fontStyleClass: string, textDecoration: string, textFontSize: number,
                         textFontFamily: string, backgroundColor: string, textTrans: string, staticTextShadow: string,
                         staticTextShadowBlur: string, staticTextShadowColor: string, fontWeight: string): void {
            this.finalTextContainer.append("span")
                .classed("staticText", true)
                .text(text)
                .classed(fontStyleClass, true)
                .style("font-size", this.pointToPixel(textFontSize))
                .style("font-family", textFontFamily)
                .style("background-color", backgroundColor +
                    this.getOpacityHex(this.staticTextSettings.transparency === null
                        ? 0 : this.staticTextSettings.transparency))
                .style("text-decoration", textDecoration)
                .style("text-shadow", this.getTextShadow(staticTextShadow, staticTextShadowBlur, staticTextShadowColor))
                .style("font-weight", fontWeight)
                .style("text-transform", textTrans)
                .style("border-radius", "5px");
        }

        private getText(text: string, fontStyleClass: string, textDecoration: string, textFontSize: number,
                        dynTextShadow: string, dynTextShadowBlur: string, dynTextShadowColor: string,
                        textFontFamily: string, backgroundColor: string, fontWeight: string, textTransD: string): void {
            this.finalTextContainer.append("span")
                .classed("dynamicText", true)
                .text(text)
                .classed(fontStyleClass, true)
                .style("font-size", this.pointToPixel(textFontSize))
                .style("font-family", textFontFamily)
                .style("text-shadow", this.getTextShadow(dynTextShadow, dynTextShadowBlur, dynTextShadowColor))
                .style("font-weight", fontWeight)
                .style("background-color", backgroundColor +
                    this.getOpacityHex(this.dynamicSettings.transparency === null
                        ? 0 : this.dynamicSettings.transparency))
                .style("text-decoration", textDecoration)
                .style("text-transform", textTransD)
                .style("border-radius", "5px");
        }

        private colonText(colonText: string): void {
            this.finalTextContainer.append("span")
                .classed("dynamicpluscolon", true)
                .text(colonText);
        }

        private addSpace(): void {
            this.finalTextContainer.append("span")
                .classed("space", true)
                .text(" ");
        }
    }
}
