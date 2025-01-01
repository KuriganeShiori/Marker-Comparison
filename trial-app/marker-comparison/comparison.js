const { DataHandler } = require('./data');

class MarkerComparison {
    static get MARKER_ORDER() {
        return [
            'D3S1358', 'vWA', 'D16S539', 'CSF1PO', 'D6S1043',
            'Yindel', 'AMEL', 'D8S1179', 'D21S11', 'D18S51',
            'D5S818', 'D2S441', 'D19S433', 'FGA', 'D10S1248',
            'D22S1045', 'D1S1656', 'D13S317', 'D7S820', 'Penta E',
            'Penta D', 'TH01', 'D12S391', 'D2S1338', 'TPOX'
        ];
    }

    constructor(dataHandler) {
        this.dataHandler = dataHandler;
    }

    async findSample(targetCode) {
        const allCases = await this.dataHandler.getAllCases();
        for (const caseData of allCases) {
            const sample = caseData.samples.find(s => s.code === targetCode);
            if (sample) return sample;
        }
        return null;
    }

    async compareFamily(baseCode) {
        const allCases = await this.dataHandler.getAllCases();
        console.log('Searching for cases with base code:', baseCode);

        const targetCase = allCases.find(c => c.baseCode === baseCode);
        if (!targetCase) {
            throw new Error(`No case found with base code ${baseCode}`);
        }

        console.log('Found case:', targetCase.baseCode, 'with samples:', 
            targetCase.samples.map(s => s.code));

        const sampleA = targetCase.samples.find(s => s.code.endsWith('A'));
        if (!sampleA) {
            throw new Error(`No sample A found for case ${baseCode}`);
        }

        const results = [];
        for (const otherSample of targetCase.samples) {
            if (otherSample.code !== sampleA.code) {
                console.log(`Comparing ${sampleA.code} with ${otherSample.code}`);
                results.push(this.compareTwoSamples(sampleA, otherSample));
            }
        }

        return results;
    }

    async compareSameDay(sampleCode) {
        const allCases = await this.dataHandler.getAllCases();
        const targetSample = await this.findSample(sampleCode);
        if (!targetSample) {
            throw new Error(`Sample ${sampleCode} not found`);
        }

        const baseCode = targetSample.code.substring(0, 5);
        const results = [];

        for (const caseData of allCases) {
            for (const sample of caseData.samples) {
                if (!sample.code.startsWith(baseCode)) {
                    results.push(this.compareTwoSamples(targetSample, sample));
                }
            }
        }
        results.sort((a, b) => {
            if (a.conclusion === b.conclusion) return 0;
            return a.conclusion === "Blood Relation" ? -1 : 1;
        });
        return results;
    }

    async compareAllDatabase(sampleCode) {
        const allCases = await this.dataHandler.getAllCases();
        const targetSample = await this.findSample(sampleCode);
        if (!targetSample) {
            throw new Error(`Sample ${sampleCode} not found`);
        }

        const results = [];
        for (const caseData of allCases) {
            for (const sample of caseData.samples) {
                if (sample.code !== targetSample.code) {
                    results.push(this.compareTwoSamples(targetSample, sample));
                }
            }
        }
        results.sort((a, b) => {
            if (a.conclusion === b.conclusion) return 0;
            return a.conclusion === "Blood Relation" ? -1 : 1;
        });
        return results;
    }

    async compareSamples(code1, code2) {
        const sample1 = await this.findSample(code1);
        const sample2 = await this.findSample(code2);

        if (!sample1 || !sample2) {
            throw new Error('One or both samples not found');
        }

        return [this.compareTwoSamples(sample1, sample2)];
    }

    compareTwoSamples(sample1, sample2) {
        if (!sample1 || !sample2) {
            console.error('Invalid samples:', { sample1, sample2 });
            throw new Error('Invalid samples provided for comparison');
        }

        if (!sample1.markers || !sample2.markers) {
            console.error('Missing markers:', {
                sample1Code: sample1.code,
                sample2Code: sample2.code,
                sample1Markers: !!sample1.markers,
                sample2Markers: !!sample2.markers
            });
            throw new Error('Missing marker data for comparison');
        }

        console.log('Comparing samples:', {
            sample1: { code: sample1.code, markers: Object.keys(sample1.markers).length },
            sample2: { code: sample2.code, markers: Object.keys(sample2.markers).length }
        });

        const matches = [];
        const mismatches = [];

        MarkerComparison.MARKER_ORDER.forEach(marker => {
            const values1 = sample1.markers[marker] || ['', ''];
            const values2 = sample2.markers[marker] || ['', ''];

            console.log(`Comparing marker ${marker}:`, { values1, values2 });

            if (this.hasMatchingAlleles(values1, values2)) {
                matches.push(marker);
            } else {
                mismatches.push(marker);
            }
        });

        console.log('Comparison result:', {
            matches: matches.length,
            mismatches: mismatches.length
        });

        const conclusion = mismatches.length <= 1 ? "Blood Relation" : "No Blood Relation";

        return {
            sample1,
            sample2,
            matches,
            mismatches,
            conclusion
        };
    }

    hasMatchingAlleles(values1, values2) {
        return values1.some(v1 => values2.includes(v1));
    }
}

module.exports = MarkerComparison; 
 