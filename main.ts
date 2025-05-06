import { App, Plugin, MarkdownPostProcessorContext, TFile } from 'obsidian';
import * as XLSX from 'xlsx';
import * as path from 'path';

interface SpreadsheetRange {
	sheet?: string;
	startCell: string;
	endCell: string;
}

export default class SpreadsheetSyncPlugin extends Plugin {
	async onload() {
		// Register the markdown processor for spreadsheet codeblocks
		this.registerMarkdownCodeBlockProcessor('spreadsheet', async (source, el, ctx) => {
			try {
				await this.renderSpreadsheetBlock(source, el, ctx);
			} catch (error) {
				el.createEl('div', { text: `Error: ${error.message}`, cls: 'spreadsheet-error' });
			}
		});
	}

	private parseRange(rangeStr: string): SpreadsheetRange {
		// Match patterns like "Sheet1!A1:B2" or "A1:B2"
		const match = rangeStr.match(/(?:([^!]+)!)?([A-Z]+\d+):([A-Z]+\d+)/);
		if (!match) {
			throw new Error('Invalid range format. Expected format: [SheetName!]A1:B2');
		}

		return {
			sheet: match[1],
			startCell: match[2],
			endCell: match[3]
		};
	}

	private async renderSpreadsheetBlock(source: string, el: HTMLElement, ctx: MarkdownPostProcessorContext) {
		// Parse the source: filename.xlsx(A1:B2) or filename.xlsx(Sheet1!A1:B2)
		const match = source.trim().match(/^(.+?)\((.+?)\)$/);
		if (!match) {
			throw new Error('Invalid format. Expected: filename.xlsx(A1:B2) or filename.xlsx(Sheet1!A1:B2)');
		}

		const [_, filename, rangeStr] = match;
		const range = this.parseRange(rangeStr);

		// Get the full path to the spreadsheet file
		const notePath = ctx.sourcePath;
		const noteDir = path.dirname(notePath);
		// Convert Windows path separators to forward slashes for Obsidian's vault API
		const filePath = path.join(noteDir, filename).replace(/\\/g, '/');
		
		console.log('SpreadsheetSync Debug:', {
			notePath,
			noteDir,
			filename,
			filePath,
			exists: this.app.vault.getAbstractFileByPath(filePath) !== null
		});

		// Read the file using the Obsidian API
		const abstractFile = this.app.vault.getAbstractFileByPath(filePath);
		if (!abstractFile || !(abstractFile instanceof TFile)) {
			throw new Error(`File not found or is not a valid file: ${filename} (looking in ${filePath})`);
		}
		const file = abstractFile;

		// Read the file contents
		const arrayBuffer = await this.app.vault.readBinary(file);
		const workbook = XLSX.read(arrayBuffer, { type: 'array' });

		// Determine which sheet to use
		const sheetName = range.sheet || workbook.SheetNames[0];
		const worksheet = workbook.Sheets[sheetName];
		if (!worksheet) {
			throw new Error(`Sheet not found: ${sheetName}`);
		}

		// Get the data from the specified range
		const rangeData = XLSX.utils.sheet_to_json(worksheet, {
			header: 1,
			range: `${range.startCell}:${range.endCell}`
		}) as any[][];

		// Create the table
		const table = el.createEl('table', { cls: 'spreadsheet-table' });
		
		// Create table rows
		rangeData.forEach((row, rowIndex) => {
			const tr = table.createEl('tr');
			row.forEach((cell) => {
				const td = tr.createEl(rowIndex === 0 ? 'th' : 'td');
				td.textContent = cell?.toString() || '';
			});
		});
	}
}
