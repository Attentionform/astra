/**
 * Single source of truth for the c:barChart 2D matrix: bar direction ×
 * grouping → chart type. The writer (manipulators/chart2d.ts) and the reader
 * (parser/factory.ts) both consume this table, so a new variant can never be
 * written without also being detected (and vice versa).
 */
export type Chart2DBarType =
  | 'column' | 'bar' | 'columnStacked' | 'barStacked'
  | 'columnPercentStacked' | 'barPercentStacked';

export interface BarChartVariant {
  barDir: 'col' | 'bar';
  grouping: 'clustered' | 'stacked' | 'percentStacked';
  percent: boolean;
}

export const BAR_2D_VARIANTS: Record<Chart2DBarType, BarChartVariant> = {
  column: { barDir: 'col', grouping: 'clustered', percent: false },
  bar: { barDir: 'bar', grouping: 'clustered', percent: false },
  columnStacked: { barDir: 'col', grouping: 'stacked', percent: false },
  barStacked: { barDir: 'bar', grouping: 'stacked', percent: false },
  columnPercentStacked: { barDir: 'col', grouping: 'percentStacked', percent: true },
  barPercentStacked: { barDir: 'bar', grouping: 'percentStacked', percent: true },
};

export function barChartVariantOf(chartType: string): BarChartVariant | undefined {
  return (BAR_2D_VARIANTS as Record<string, BarChartVariant | undefined>)[chartType];
}
