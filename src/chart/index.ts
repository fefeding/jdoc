/**
 * 图表共享模块
 *
 * 各格式（PPTX / XLSX）只负责把 OOXML 解析成 `ChartSpec`，
 * 画法统一由 `chart/svg` 提供，避免每个格式各写一套 SVG 渲染器。
 *
 * @module chart/index
 */

export {
    renderChartSvg,
    DEFAULT_CHART_PALETTE,
    type ChartType,
    type ChartLegendPos,
    type ChartSeriesSpec,
    type ChartLabels,
    type ChartSpec,
    type ChartRenderOptions
} from './svg';
