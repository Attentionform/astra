/**
 * The kinds of bar and column chart: a direction (`column` is vertical and `bar` is horizontal), and grouping (clustered, stacked or percent stacked).
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
