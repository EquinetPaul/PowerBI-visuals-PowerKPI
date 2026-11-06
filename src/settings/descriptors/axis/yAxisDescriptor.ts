
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
import ILocalizationManager = powerbi.extensibility.ILocalizationManager;

import { IDescriptorParserOptions } from "../baseDescriptor";
import { AxisDescriptor } from "./axisDescriptor";

export enum FixedLabelWidthMode {
    auto = "auto",
    fixed = "fixed",
}

const fixedLabelWidthModeOptions = [
    {
        value: FixedLabelWidthMode.auto,
        displayName: "Auto",
        displayNameKey: "Visual_Auto"
    },
    {
        value: FixedLabelWidthMode.fixed,
        displayName: "Fixed",
        displayNameKey: "Visual_Fixed"
    }
]

/**
 * How the series bound to an axis are drawn relative to each other. Standard draws every
 * series from its own value, Stacked piles them up so that the top of the last one is the
 * total, and Stacked 100% scales each pile to the share of the total it represents.
 */
export enum YAxisDisplayMode {
    standard = "standard",
    stacked = "stacked",
    stacked100 = "stacked100",
}

const displayModeOptions = [
    {
        value: YAxisDisplayMode.standard,
        displayName: "Standard",
        displayNameKey: "Visual_Display_Mode_Standard"
    },
    {
        value: YAxisDisplayMode.stacked,
        displayName: "Stacked",
        displayNameKey: "Visual_Display_Mode_Stacked"
    },
    {
        value: YAxisDisplayMode.stacked100,
        displayName: "Stacked 100%",
        displayNameKey: "Visual_Display_Mode_Stacked_100"
    }
]

export class YAxisDescriptor extends AxisDescriptor {
    public displayMode = new formattingSettings.ItemDropdown({
        name: "displayMode",
        displayNameKey: "Visual_Display_Mode",
        descriptionKey: "Visual_Display_Mode_Description",
        items: displayModeOptions,
        value: displayModeOptions.filter(el => el.value === YAxisDisplayMode.standard)[0]
    });

    public min = new formattingSettings.NumUpDown({
        name: "min",
        displayNameKey: "Visual_Min",
        value: NaN,
    });
    public max = new formattingSettings.NumUpDown({
        name: "max",
        displayNameKey: "Visual_Max",
        value: NaN,
    });

    public fixedLabelWidthMode = new formattingSettings.ItemDropdown({
        name: "fixedLabelWidthMode",
        displayNameKey: "Visual_Label_Area_Sizing",
        descriptionKey: "Visual_Label_Area_Sizing_Description",
        items: fixedLabelWidthModeOptions,
        value: fixedLabelWidthModeOptions.filter(el => el.value === FixedLabelWidthMode.auto)[0]
    });

    protected minFixedLabelWidth: number = 0;
    protected maxFixedLabelWidth: number = 400;
    public fixedLabelWidth = new formattingSettings.NumUpDown({
        name: "fixedLabelWidth",
        displayNameKey: "Visual_Label_Area_Width",
        value: 60,
        options: {
            minValue: {
                type: powerbi.visuals.ValidatorType.Min,
                value: this.minFixedLabelWidth,
            },
            maxValue: {
                type: powerbi.visuals.ValidatorType.Max,
                value: this.maxFixedLabelWidth,
            },
        }
    });

    constructor(
        name: string, 
        displayNameKey: string,
        viewportToBeHidden: powerbi.IViewport,
        viewportToIncreaseDensity: powerbi.IViewport,
    ) {
        super(viewportToBeHidden, viewportToIncreaseDensity)

        this.slices = [
            this.displayMode,
            this.font,
            this.displayUnits,
            this.precision,
            this.fontColor,
            this.percentile,
            this.min,
            this.max,
            this.fixedLabelWidthMode,
            this.fixedLabelWidth
        ]
        this.name = name;
        this.displayNameKey = displayNameKey;
    }

    public parse(options: IDescriptorParserOptions) {
        super.parse(options);

        // The formatting pane validators only guard UI input, so values persisted by
        // older reports or set through the API are clamped here, the same way
        // NumberDescriptorBase clamps precision
        if (this.fixedLabelWidth.value < this.fixedLabelWidth.options.minValue.value) {
            this.fixedLabelWidth.value = this.fixedLabelWidth.options.minValue.value;
        }
        if (this.fixedLabelWidth.value > this.fixedLabelWidth.options.maxValue.value) {
            this.fixedLabelWidth.value = this.fixedLabelWidth.options.maxValue.value;
        }
    }

    public getDisplayMode(): YAxisDisplayMode {
        const mode: YAxisDisplayMode = this.displayMode.value?.value as YAxisDisplayMode;
        return mode === YAxisDisplayMode.stacked || mode === YAxisDisplayMode.stacked100
            ? mode
            : YAxisDisplayMode.standard;
    }

    public isStacked(): boolean {
        return this.getDisplayMode() !== YAxisDisplayMode.standard;
    }

    public isStacked100(): boolean {
        return this.getDisplayMode() === YAxisDisplayMode.stacked100;
    }

    public isLabelWidthFixed(): boolean {
        return this.fixedLabelWidthMode.value.value === FixedLabelWidthMode.fixed;
    }

    public setLocalizedDisplayName(localizationManager: ILocalizationManager) {
        super.setLocalizedDisplayName(localizationManager);
        [...fixedLabelWidthModeOptions, ...displayModeOptions].forEach(option => {
            option.displayName = localizationManager.getDisplayName(option.displayNameKey)
        });
    }
}
