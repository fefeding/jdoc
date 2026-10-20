import { describe, expect, it } from 'vitest';
import { pptxChartToSpec } from '../src/pptx/chart';
import { renderChartSvg } from '../src/chart/index';

describe('pptx 图表 → 共享渲染模型', () => {
    it('chartType 映射到共享类型（3D / 特殊类型降级为同族 2D）', () => {
        const cases: [string, string][] = [
            ['barChart', 'column'],
            ['barChart|bar', 'bar'],
            ['lineChart', 'line'],
            ['areaChart', 'area'],
            ['pieChart', 'pie'],
            ['doughnutChart', 'doughnut'],
            ['scatterChart', 'scatter'],
            ['radarChart', 'radar'],
            ['pie3DChart', 'pie'],
            ['surfaceChart', 'line']
        ];
        for (const [raw, type] of cases) {
            const spec = pptxChartToSpec({ chartType: raw, series: [] });
            expect(spec.type, raw).toBe(type);
        }
    });

    it('系列取值兼容 {y} / {x,y} / 裸数字', () => {
        const spec = pptxChartToSpec({
            chartType: 'scatterChart',
            categories: [],
            series: [
                { name: 'a', values: [{ x: 1, y: 2 }] },
                { name: 'b', values: [3] }
            ]
        });
        expect(spec.series?.[0].values).toEqual([2]);
        expect(spec.series?.[0].xValues).toEqual([1]);
        expect(spec.series?.[1].values).toEqual([3]);
        expect(spec.series?.[1].xValues).toBeUndefined();
    });

    it('图例位置、数据标签、分组、内径、平滑都映射过来', () => {
        const spec = pptxChartToSpec({
            chartType: 'barChart',
            legend: true,
            legendPosition: 'b',
            dataLabels: true,
            grouping: 'stacked',
            holeSize: 40,
            smooth: true,
            marker: false,
            varyColors: true,
            title: '标题',
            categories: ['A'],
            series: [{ name: 's', values: [1] }]
        });
        expect(spec.legend).toBe('b');
        expect(spec.labels).toEqual({ value: true });
        expect(spec.grouping).toBe('stacked');
        expect(spec.holeSize).toBe(40);
        expect(spec.smooth).toBe(true);
        expect(spec.marker).toBe(false);
        expect(spec.varyColors).toBe(true);
        expect(spec.title).toBe('标题');
    });

    it('主题色覆盖到 spec，且能直接渲染出 SVG', () => {
        const spec = pptxChartToSpec(
            { chartType: 'pieChart', categories: ['A'], series: [{ name: 's', values: [1] }] },
            { palette: ['#123456'], textColor: '#333333', titleColor: '#111111' }
        );
        expect(spec.palette).toEqual(['#123456']);
        expect(spec.textColor).toBe('#333333');
        const svg = renderChartSvg(spec, { width: 200, height: 150 });
        expect(svg).toContain('#123456');
    });
});
