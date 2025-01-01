const { Document, Packer } = require('docx');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const fs = require('fs');
const path = require('path');
const MarkerComparison = require('./comparison');

class DocGenerator {
    static getMarkerColor(index) {
        if (index < 5) return '0000FF'; // Blue
        if (index < 11) return '008000'; // Green
        if (index < 15) return '000000'; // Black
        if (index < 20) return 'FF0000'; // Red
        return '800080'; // Purple
    }

    static async generateDocument(result) {
        try {
            const template = fs.readFileSync(
                path.resolve(__dirname, 'templates', 'phap-ly-template.docx'),
                'binary'
            );

            const zip = new PizZip(template);
            
            const doc = new Docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
                nullGetter() { return ''; },
                delimiters: {
                    start: '{{',
                    end: '}}'
                }
            });

            // Get the marker order array
            const markerOrder = MarkerComparison.MARKER_ORDER;
            if (!markerOrder || !Array.isArray(markerOrder)) {
                throw new Error('MARKER_ORDER is not properly defined');
            }

            // Check if there are mismatches excluding Yindel
            const hasMismatchesExcludingYindel = result.mismatches.some(marker => marker !== 'Yindel');
            const showPercentage = !hasMismatchesExcludingYindel && result.conclusion === "Blood Relation";

            // Create data structure for template
            const data = {
                sample1_name: result.sample1.name || '',
                sample1_code: result.sample1.code || '',
                sample2_name: result.sample2.name || '',
                sample2_code: result.sample2.code || '',
                Conclusion: result.conclusion === "Blood Relation" ? "CÓ" : "KHÔNG",
                showPercentage: showPercentage,
                percentage: showPercentage ? "99.99999999999999" : ""
            };

            // Generate marker data for template
            markerOrder.forEach((marker, index) => {
                const values1 = result.sample1.markers[marker] || ['', ''];
                const values2 = result.sample2.markers[marker] || ['', ''];
                const safeMarker = marker.replace(/\s+/g, '_');

                data[`${safeMarker}_values1_0`] = values1[0] || '';
                data[`${safeMarker}_values1_1`] = values1[1] || '';
                data[`${safeMarker}_values2_0`] = values2[0] || '';
                data[`${safeMarker}_values2_1`] = values2[1] || '';
            });

            // Debug logging
            console.log('Template path:', path.resolve(__dirname, 'templates', 'phap-ly-template.docx'));
            console.log('Data being sent to template:', JSON.stringify(data, null, 2));

            // Render the document
            doc.render(data);

            return doc.getZip().generate({
                type: 'nodebuffer',
                compression: 'DEFLATE'
            });

        } catch (error) {
            console.error('Original error:', error);
            if (error.properties && error.properties.errors) {
                const errorMessages = error.properties.errors
                    .map(e => `Error in ${e.properties?.file || 'template'}: ${e.properties?.explanation || e.message}`)
                    .join('\n');
                console.error('Template Error Details:', errorMessages);
                throw new Error(`Template processing failed: ${errorMessages}`);
            }
            throw error;
        }
    }

    static async downloadDocument(result) {
        try {
            const buffer = await this.generateDocument(result);

            // Convert buffer to blob
            const blob = new Blob([buffer], { 
                type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
            });
            
            // Format filename
            let filename;
            if (result.sample1.code.endsWith('A')) {
                filename = `${result.sample1.code} ${result.sample1.name}_${result.sample2.name}`;
            } else if (result.sample2.code.endsWith('A')) {
                filename = `${result.sample2.code} ${result.sample2.name}_${result.sample1.name}`;
            } else {
                filename = `${result.sample1.code} ${result.sample1.name}_${result.sample2.name}`;
            }

            filename = filename.replace(/[/\\?%*:|"<>]/g, '-');
            filename = `${filename}.docx`;

            // Download file
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            document.body.appendChild(a);
            a.style.display = 'none';
            a.href = url;
            a.download = filename;
            a.click();
            window.URL.revokeObjectURL(url);
        } catch (error) {
            console.error('Error generating document:', error);
            alert('Error generating document: ' + error.message);
        }
    }
}

module.exports = { DocGenerator }; 