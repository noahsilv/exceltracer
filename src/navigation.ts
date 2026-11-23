import type { CellRef, Mode, Trace } from './excel.js';

export class Navigator {
  current?: Trace;
  mode: Mode = 'inputs';
  private trail: Array<{ cell: CellRef; mode: Mode; via: CellRef }> = [];
  constructor(private inspect: (mode: Mode, target?: CellRef, selectTarget?: boolean) => Promise<Trace>) {}
  get canGoBack() { return this.trail.length > 0; }

  async refresh(mode: Mode) {
    const next = await this.inspect(mode);
    if (this.current?.sheetId !== next.sheetId || this.current?.address !== next.address) this.trail = [];
    this.current = next;
    this.mode = mode;
  }

  async jump(target: CellRef, mode: Mode = this.mode) {
    const next = await this.inspect(mode, target);
    if (this.current) {
      const { sheetId, address } = this.current;
      this.trail.push({ cell: { sheetId, address }, mode: this.mode, via: {sheetId:target.sheetId, address:target.address} });
      if (this.trail.length > 100) this.trail.shift();
    }
    this.current = next;
    this.mode = mode;
  }

  async back(selectRoot = true) {
    const prior = this.trail[this.trail.length - 1];
    if (!prior) return;
    const next = await this.inspect(prior.mode, prior.cell, selectRoot);
    this.trail.pop();
    this.current = next;
    this.mode = prior.mode;
    return prior.via;
  }
}
