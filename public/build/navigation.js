export class Navigator {
    constructor(inspect) {
        this.inspect = inspect;
        this.mode = 'inputs';
        this.trail = [];
    }
    get canGoBack() { return this.trail.length > 0; }
    async refresh(mode) {
        const next = await this.inspect(mode);
        if (this.current?.sheetId !== next.sheetId || this.current?.address !== next.address)
            this.trail = [];
        this.current = next;
        this.mode = mode;
    }
    async jump(target, mode = this.mode) {
        const next = await this.inspect(mode, target);
        if (this.current) {
            const { sheetId, address } = this.current;
            this.trail.push({ cell: { sheetId, address }, mode: this.mode, via: { sheetId: target.sheetId, address: target.address } });
            if (this.trail.length > 100)
                this.trail.shift();
        }
        this.current = next;
        this.mode = mode;
    }
    async back(selectRoot = true) {
        const prior = this.trail[this.trail.length - 1];
        if (!prior)
            return;
        const next = await this.inspect(prior.mode, prior.cell, selectRoot);
        this.trail.pop();
        this.current = next;
        this.mode = prior.mode;
        return prior.via;
    }
}
