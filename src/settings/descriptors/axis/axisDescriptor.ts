/**
 *  Power BI Visualizations
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

import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import { pixelConverter } from "powerbi-visuals-utils-typeutils";

import {
    IDescriptor,
    IDescriptorParserOptions,
} from "../baseDescriptor";
import { NumberDescriptorBase } from "../numberDescriptorBase";

export enum AxisType {
    continuous,
    categorical,
}

/**
 * What an axis title is drawn with, resolved from the settings of its card.
 */
export interface IAxisTitle {
    text: string;
    fontSizeInPx: number;
    fontFamily: string;
    color: string;
    isBold: boolean;
    isItalic: boolean;
}

export class AxisDescriptor
    extends NumberDescriptorBase
    implements IDescriptor {

    public maxDensity: number = 100;

    public fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayNameKey: "Visual_Font_Color",
        value: { value: "rgb(0,0,0)" }
    });

    // Off by default, so a report built before axis titles existed renders unchanged
    public showTitle = new formattingSettings.ToggleSwitch({
        name: "showTitle",
        displayNameKey: "Visual_Axis_Title",
        value: false
    });

    // Left empty, the title names what the axis shows: the field of the X axis, the
    // measures of a Y axis
    public titleText = new formattingSettings.TextInput({
        name: "titleText",
        displayNameKey: "Visual_Title_Text",
        value: "",
        placeholder: ""
    });

    protected minTitleFontSize: number = 8;
    protected maxTitleFontSize: number = 60;

    public titleFontSize = new formattingSettings.NumUpDown({
        name: "titleFontSize",
        displayNameKey: "Visual_Font_Size",
        value: 9,
        options: {
            minValue: {
                type: powerbi.visuals.ValidatorType.Min,
                value: this.minTitleFontSize,
            },
            maxValue: {
                type: powerbi.visuals.ValidatorType.Max,
                value: this.maxTitleFontSize,
            },
        }
    });

    public titleFontFamily = new formattingSettings.FontPicker({
        name: "titleFontFamily",
        displayNameKey: "Visual_Font_Family",
        value: "Segoe UI, wf_segoe-ui_normal, helvetica, arial, sans-serif"
    });

    public titleBold = new formattingSettings.ToggleSwitch({
        name: "titleBold",
        value: false
    });

    public titleItalic = new formattingSettings.ToggleSwitch({
        name: "titleItalic",
        value: false
    });

    public titleFont = new formattingSettings.FontControl({
        name: "titleFont",
        displayNameKey: "Visual_Title_Font",
        fontFamily: this.titleFontFamily,
        fontSize: this.titleFontSize,
        bold: this.titleBold,
        italic: this.titleItalic,
    });

    public titleColor = new formattingSettings.ColorPicker({
        name: "titleColor",
        displayNameKey: "Visual_Title_Color",
        value: { value: "rgb(102,102,102)" }
    });

    private shouldDensityBeAtMax: boolean = false;
    private viewportToIncreaseDensity: powerbi.IViewport;

    constructor(
        viewportToBeHidden: powerbi.IViewport,
        viewportToIncreaseDensity: powerbi.IViewport,
        shouldPropertiesBeHiddenByType: boolean = false
    ) {
        super(viewportToBeHidden, shouldPropertiesBeHiddenByType);

        this.viewportToIncreaseDensity = viewportToIncreaseDensity;
        this.font.fontFamily.value = "Segoe UI Light, wf_segoe-ui_light, helvetica, arial, sans-serif"
        this.percentile.options.maxValue.value = this.maxDensity
    }

    public parse(options: IDescriptorParserOptions) {
        super.parse(options);

        // The formatting pane validators only guard UI input, so a value set through the
        // API is clamped here as well
        this.titleFontSize.value = Math.min(
            this.maxTitleFontSize,
            Math.max(this.minTitleFontSize, this.titleFontSize.value || this.minTitleFontSize),
        );

        if (this.shouldDensityBeAtMax) {
            this.percentile.value = this.maxDensity;
        }

        this.shouldDensityBeAtMax = options.isAutoHideBehaviorEnabled
            && this.viewportToIncreaseDensity
            && options.viewport
            && (options.viewport.width <= this.viewportToIncreaseDensity.width
                ||
                options.viewport.height <= this.viewportToIncreaseDensity.height);
    }

    /**
     * The settings of the title that follow the Title toggle, shown in the pane only
     * while the title is on.
     */
    public getTitleSlices(): formattingSettings.Slice[] {
        return [this.titleText, this.titleFont, this.titleColor];
    }

    /**
     * The title to draw next to the axis, or null when there is none: the axis or its
     * title is turned off, or there is no text to show.
     *
     * @param defaultText What the title reads when the author left its text empty.
     */
    public getTitle(defaultText: string): IAxisTitle | null {
        if (!this.isElementShown() || !this.showTitle.value) {
            return null;
        }

        const text: string = (this.titleText.value || "").trim() || (defaultText || "").trim();

        if (!text) {
            return null;
        }

        return {
            color: this.titleColor.value.value,
            fontFamily: this.titleFontFamily.value,
            fontSizeInPx: pixelConverter.fromPointToPixel(this.titleFontSize.value),
            isBold: !!this.titleBold.value,
            isItalic: !!this.titleItalic.value,
            text,
        };
    }
}
