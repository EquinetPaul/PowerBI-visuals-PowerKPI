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
import "../styles/styles.less";
import {
    select as d3Select,
    Selection,
} from "d3-selection";
import { dispatch, Dispatch } from "d3-dispatch";

import powerbi from "powerbi-visuals-api";
import {
    interactivityBaseService,
    interactivitySelectionService,
} from "powerbi-visuals-utils-interactivityutils";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

import { IConverter } from "./converter/converter";
import { DataConverter } from "./converter/dataConverter";
import { IDataRepresentation } from "./dataRepresentation/dataRepresentation";
import { EventName } from "./event/eventName";
import {
    createDefaultZoom,
    IDataRepresentationZoom,
    IDataRepresentationZoomRange,
} from "./dataRepresentation/dataRepresentationZoom";
import { ZoomSliderAxis } from "./visualComponent/zoomSliderComponent";
import { IVisualComponent } from "./visualComponent/base/visualComponent";
import { IVisualComponentRenderOptions } from "./visualComponent/base/visualComponentRenderOptions";
import { MainComponent } from "./visualComponent/mainComponent";

import { IDataRepresentationSeries } from "./dataRepresentation/dataRepresentationSeries";

import {
    Behavior,
    IBehaviorOptions,
} from "./behavior/behavior";
import { Settings } from "./settings/settings";
import { DataRepresentationTypeEnum } from "./dataRepresentation/dataRepresentationType";
import { AxisType } from "./settings/descriptors/axis/axisDescriptor";
import { NumberDescriptorBase } from "./settings/descriptors/numberDescriptorBase";

import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import ILocalizationManager = powerbi.extensibility.ILocalizationManager;
import ISandboxExtendedColorPalette = powerbi.extensibility.ISandboxExtendedColorPalette;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

export interface IPowerKPIConstructorOptions extends VisualConstructorOptions {
    rootElement?: HTMLElement;
}

export class PowerKPI implements IVisual {
    private static ViewportReducer: number = 3;

    private axisType: DataRepresentationTypeEnum = DataRepresentationTypeEnum.None;
    private eventDispatcher: Dispatch<any> = dispatch(...Object.keys(EventName));
    private element: Selection<any, any, any, any>;
    private converter: IConverter;
    private component: IVisualComponent<IVisualComponentRenderOptions>;
    private dataRepresentation: IDataRepresentation;
    private behavior: Behavior;
    private interactivityService: interactivityBaseService.IInteractivityService<IDataRepresentationSeries>;
    private events: IVisualEventService;

    private localizationManager: ILocalizationManager;
    private settings: Settings;
    private zoom: IDataRepresentationZoom = createDefaultZoom();
    // The fields the zoom was picked on, see getBoundFieldsKey
    private zoomFieldsKey: string;
    private updateOptions: VisualUpdateOptions;
    private formattingSettingsService: FormattingSettingsService;
    private colorPalette: ISandboxExtendedColorPalette;
    private host: IVisualHost;

    constructor(options: IPowerKPIConstructorOptions) {
        this.events = options.host.eventService;
        this.settings = new Settings()
        this.localizationManager = options.host.createLocalizationManager()
        this.formattingSettingsService = new FormattingSettingsService(this.localizationManager);
        this.element = d3Select(options.element);
        this.colorPalette = options.host.colorPalette;

        this.host = options.host;

        this.converter = new DataConverter({
            colorPalette: options.host.colorPalette,
            createSelectionIdBuilder: options.host.createSelectionIdBuilder.bind(options.host),
        });

        this.behavior = new Behavior();
        this.interactivityService = interactivitySelectionService.createInteractivitySelectionService(options.host);

        const rootElement = options.rootElement && d3Select(options.rootElement) || this.element;

        this.component = new MainComponent({
            element: this.element,
            eventDispatcher: this.eventDispatcher,
            interactivityService: this.interactivityService,
            rootElement,
            tooltipService: options.host.tooltipService,
        });

        this.eventDispatcher.on(
            EventName.onZoom,
            (axis: ZoomSliderAxis, range: IDataRepresentationZoomRange) => this.applyZoom(axis, range),
        );
    }

    public update(options: VisualUpdateOptions): void {
        this.events.renderingStarted(options);

        // The zoom is view state that has to survive every update that keeps the same
        // fields - a resize, a change of view mode, a formatting change, a filter - just
        // like the zoom of the built in visuals. Being a fraction of the axis range, it
        // still means the same thing once a filter narrows the data. It is only reset
        // when the fields bound to the visual change, a range picked on other fields
        // having no meaning on the new ones.
        const fieldsKey: string = PowerKPI.getBoundFieldsKey(options && options.dataViews && options.dataViews[0]);

        if (fieldsKey !== this.zoomFieldsKey) {
            this.zoom = createDefaultZoom();
            this.zoomFieldsKey = fieldsKey;
        }

        this.updateOptions = options;

        this.convertAndRender(options);

        this.events.renderingFinished(options);
    }

    /**
     * Rebuilds the data representation and renders it. The viewport is reduced in
     * place while a render walks the component tree, so a re-render has to start from
     * a freshly converted representation rather than from the previous one.
     */
    private convertAndRender(options: VisualUpdateOptions): void {
        this.settings = this.formattingSettingsService.populateFormattingSettingsModel(Settings, options.dataViews[0]);
        this.zoom = this.getZoomOfShownSliders(this.zoom, this.settings);

        const dataView: powerbi.DataView = options && options.dataViews && options.dataViews[0];

        const viewport: powerbi.IViewport = options
            && options.viewport
            && {
                height: options.viewport.height - PowerKPI.ViewportReducer,
                width: options.viewport.width - PowerKPI.ViewportReducer,
            }
            || { height: 0, width: 0 };

        this.axisType = this.converter.getAxisType({
            dataView,
            xAxisType: this.settings.xAxis.type.value.value as AxisType
        });
        this.updateFormatPropertyValue();

        const dataRepresentation: IDataRepresentation = this.converter.convert({
            dataView,
            hasSelection: this.interactivityService && this.interactivityService.hasSelection(),
            viewport,
            settings: this.settings,
            locale: this.host.locale,
            zoom: this.zoom
        });
        if (this.interactivityService) {
            this.interactivityService.applySelectionStateToData(dataRepresentation.series);

            const behaviorOptions: IBehaviorOptions = {
                behavior: this.behavior,
                dataPoints: dataRepresentation.series,
                eventDispatcher: this.eventDispatcher,
                interactivityService: this.interactivityService,
            };

            this.interactivityService.bind(behaviorOptions);
        }

        this.render(dataRepresentation);
    }

    /**
     * Identifies the fields bound to the visual - the axis column and the measures, with
     * the role each one fills - regardless of the values a filter leaves in them.
     */
    public static getBoundFieldsKey(dataView: powerbi.DataView): string {
        const columns: powerbi.DataViewMetadataColumn[] = dataView
            && dataView.metadata
            && dataView.metadata.columns
            || [];

        return columns
            .map((column: powerbi.DataViewMetadataColumn) => {
                const roles: string = Object.keys(column.roles || {})
                    .filter((role: string) => column.roles[role])
                    .sort()
                    .join("+");

                return `${roles}:${column.queryName || column.displayName}`;
            })
            .sort()
            .join("|");
    }

    /**
     * A zoom slider that is turned off no longer narrows its axis: the range it kept would
     * otherwise stay applied with no way to see or change it.
     */
    private getZoomOfShownSliders(zoom: IDataRepresentationZoom, settings: Settings): IDataRepresentationZoom {
        const fullZoom: IDataRepresentationZoom = createDefaultZoom();

        return {
            x: settings.zoomSlider.isShownForXAxis() ? zoom.x : fullZoom.x,
            y: settings.zoomSlider.isShownForYAxis() ? zoom.y : fullZoom.y,
        };
    }

    private applyZoom(axis: ZoomSliderAxis, range: IDataRepresentationZoomRange): void {
        if (!this.updateOptions) {
            return;
        }

        this.zoom = {
            ...this.zoom,
            [axis]: range,
        };

        this.convertAndRender(this.updateOptions);
    }

    public render(dataRepresentation: IDataRepresentation): void {
        this.dataRepresentation = dataRepresentation;

        this.component.render({
            data: this.dataRepresentation,
            colorPalette: this.colorPalette
        });
    }

    public destroy(): void {
        this.component.destroy();

        this.converter = null;
        this.element = null;
        this.component = null;
        this.dataRepresentation = null;
        this.interactivityService = null;
        this.behavior = null;
        this.settings = null;
    }

    public updateFormatPropertyValue(){
        this.settings.cards.forEach(card => {
            if(card instanceof NumberDescriptorBase){
                card.applyDefaultFormatByType(this.axisType)
            }
        })
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        this.settings.filterFormattingProperties(this.dataRepresentation, this.axisType, this.localizationManager, this.colorPalette.isHighContrast);

        return this.formattingSettingsService.buildFormattingModel(this.settings);
    }    
}
