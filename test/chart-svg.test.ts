import { describe, expect, it } from 'vitest';
import { renderChartSvg, type ChartSpec } from '../src/chart/index';

function spec(over: Partial<ChartSpec> = {}): ChartSpec {
    return {
        type: 'column',
        categories: ['A', 'B', 'C'],
        series: [{ name: '系列1', values: [1, 2, 3] }],
        ...over
    };
}

describe('共享图表渲染（chart/svg）', () => {
    it('柱状图：出方块、轴刻度与分类标签', () => {
        const svg = renderChartSvg(spec(), { width: 400, height: 240 });
        expect(svg.startsWith('<svg')).toBe(true);
        expect(svg).toContain('<rect');
        expect(svg).toContain('>A<');
        expect(svg).toContain('>B<');
    });

    it('数值格式作用于轴标签（Excel 数字格式串）', () => {
        const svg = renderChartSvg(
            spec({ series: [{ name: 's', values: [1250000, 2000000, 3000000] }], numFmt: '¥#,##0' }),
            { width: 400, height: 240 }
        );
        // 轴上限取整到 5,000,000（1/2/5 刻度）
        expect(svg).toContain('¥5,000,000');
    });

    it('饼图：扇区 + 百分比标签 + 分类图例', () => {
        const svg = renderChartSvg(
            spec({
                type: 'pie',
                categories: ['手机', '平板'],
                series: [{ name: 's', values: [75, 25] }],
                labels: { percent: true },
                legend: 'r'
            }),
            { width: 400, height: 240 }
        );
        expect(svg).toContain('75%');
        expect(svg).toContain('手机');
        expect(svg).toContain('<path');
    });

    it('环形图：内径来自 holeSize', () => {
        const svg = renderChartSvg(
            spec({ type: 'doughnut', holeSize: 40, series: [{ name: 's', values: [1, 1] }] }),
            { width: 300, height: 200 }
        );
        // 内圈是独立一段圆弧，半径小于外圈
        expect(svg).toMatch(/A\d+\.\d+,\d+\.\d+ 0 \d 0/);
    });

    it('折线图：折线路径 + 可选平滑 + 数据点标记', () => {
        const plain = renderChartSvg(spec({ type: 'line' }), { width: 400, height: 240 });
        expect(plain).toContain('<path d="M');
        expect(plain).not.toContain(' C');
        const smooth = renderChartSvg(spec({ type: 'line', smooth: true }), { width: 400, height: 240 });
        expect(smooth).toContain(' C');
        const noMarker = renderChartSvg(spec({ type: 'line', marker: false }), { width: 400, height: 240 });
        expect(noMarker).not.toContain('<circle');
    });

    it('散点图：x 用系列的 xValues 而不是序号', () => {
        const svg = renderChartSvg(
            spec({ type: 'scatter', categories: [], series: [{ name: 's', values: [1, 2], xValues: [10, 20] }] }),
            { width: 400, height: 240 }
        );
        const xs = Array.from(svg.matchAll(/<circle cx="([\d.]+)"/g)).map((m) => Number(m[1]));
        expect(xs.length).toBe(2);
        expect(xs[1] - xs[0]).toBeGreaterThan(100);
    });

    it('雷达图：画多边形网格与系列', () => {
        const svg = renderChartSvg(
            spec({ type: 'radar', categories: ['A', 'B', 'C'], series: [{ name: 's', values: [1, 2, 3] }] }),
            { width: 400, height: 240 }
        );
        expect(svg).toContain('<polygon');
    });
});

describe('共享图表渲染（分组 / 图例 / 标签）', () => {
    it('堆积图：轴上限是各系列合计', () => {
        const svg = renderChartSvg(
            spec({
                grouping: 'stacked',
                series: [
                    { name: 'a', values: [1, 2] },
                    { name: 'b', values: [1, 2] }
                ]
            }),
            { width: 400, height: 240 }
        );
        expect(svg).toContain('>3<');
    });

    it('百分比堆积：轴上限恒为 100', () => {
        const svg = renderChartSvg(
            spec({ grouping: 'percentStacked', series: [{ name: 'a', values: [1, 1] }] }),
            { width: 400, height: 240 }
        );
        expect(svg).toContain('>100<');
    });

    it('图例四个位置都能画出来', () => {
        for (const pos of ['l', 'r', 't', 'b'] as const) {
            const svg = renderChartSvg(spec({ legend: pos }), { width: 400, height: 240 });
            expect(svg, pos).toContain('系列1');
        }
    });

    it('不显示图例时不出图例文字', () => {
        const svg = renderChartSvg(spec({ legend: 'none' }), { width: 400, height: 240 });
        expect(svg).not.toContain('系列1');
    });

    it('数据标签：数值 / 分类名 / 百分比', () => {
        const svg = renderChartSvg(
            spec({
                type: 'pie',
                categories: ['手机'],
                series: [{ name: 's', values: [10] }],
                labels: { value: true, percent: true, category: true }
            }),
            { width: 400, height: 240 }
        );
        expect(svg).toContain('100%');
        expect(svg).toContain('10');
        expect(svg).toContain('手机');
    });

    it('空数据给出提示而不是空图', () => {
        const svg = renderChartSvg(
            spec({ type: 'pie', categories: ['A'], series: [{ name: 's', values: [0] }] }),
            { width: 300, height: 200 }
        );
        expect(svg).toContain('无数据');
    });
});
