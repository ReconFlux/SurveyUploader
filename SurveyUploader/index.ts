import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as XLSX from "xlsx";

interface ExcelRow {
  ID: string;
  Response: string;
  Notes: string;
}

interface UpdateResult {
  success: boolean;
  recordId: string;
  error?: string;
}

export class SurveyUploader
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  private _context: ComponentFramework.Context<IInputs>;
  private _container: HTMLDivElement;
  private _fileInput: HTMLInputElement;
  private _uploadButton: HTMLButtonElement;
  private _processButton: HTMLButtonElement;
  private _statusDiv: HTMLDivElement;
  private _progressDiv: HTMLDivElement;
  private _resultsDiv: HTMLDivElement;
  private _notifyOutputChanged: () => void;
  private _value: string;
  private _parsedData: ExcelRow[] = [];
  private _systemId: string;

  constructor() {
    this._value = "";
  }

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this._context = context;
    this._notifyOutputChanged = notifyOutputChanged;
    this._container = container;

    // Get system ID from the current form context
    this._systemId = (context as any).page?.entityId || "";

    this.createUI();
  }

  private createUI(): void {
    this._container.innerHTML = `
            <div class="excel-uploader-container">
                <div class="upload-section">
                    <h3>Upload Excel File</h3>
                    <input type="file" accept=".xlsx,.xls" class="file-input" />
                    <button class="upload-btn" disabled>Upload & Parse</button>
                </div>
                
                <div class="status-section">
                    <div class="status-message"></div>
                    <div class="progress-bar" style="display: none;">
                        <div class="progress-fill"></div>
                        <span class="progress-text">0%</span>
                    </div>
                </div>
                
                <div class="preview-section" style="display: none;">
                    <h3>Data Preview</h3>
                    <div class="data-preview"></div>
                    <button class="process-btn">Update Dataverse Records</button>
                </div>
                
                <div class="results-section" style="display: none;">
                    <h3>Update Results</h3>
                    <div class="results-content"></div>
                </div>
            </div>
        `;

    this._fileInput = this._container.querySelector(
      ".file-input"
    ) as HTMLInputElement;
    this._uploadButton = this._container.querySelector(
      ".upload-btn"
    ) as HTMLButtonElement;
    this._processButton = this._container.querySelector(
      ".process-btn"
    ) as HTMLButtonElement;
    this._statusDiv = this._container.querySelector(
      ".status-message"
    ) as HTMLDivElement;
    this._progressDiv = this._container.querySelector(
      ".progress-bar"
    ) as HTMLDivElement;
    this._resultsDiv = this._container.querySelector(
      ".results-content"
    ) as HTMLDivElement;

    this.attachEventListeners();
  }

  private attachEventListeners(): void {
    this._fileInput.addEventListener("change", (event) => {
      const file = (event.target as HTMLInputElement).files?.[0];
      this._uploadButton.disabled = !file;

      if (file) {
        this.showStatus(`Selected: ${file.name}`, "info");
      }
    });

    this._uploadButton.addEventListener("click", () => {
      this.uploadAndParseFile();
    });

    this._processButton.addEventListener("click", () => {
      this.processDataverseUpdates();
    });
  }

  private async uploadAndParseFile(): Promise<void> {
    const file = this._fileInput.files?.[0];
    if (!file) {
      this.showStatus("Please select a file first", "error");
      return;
    }

    this.showStatus("Parsing Excel file...", "info");
    this.showProgress(0);

    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: "array" });

      // Get the first sheet
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      // Convert to JSON
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      this.showProgress(50);

      // Parse the data
      this._parsedData = this.parseExcelData(jsonData);

      this.showProgress(100);
      this.showStatus(
        `Successfully parsed ${this._parsedData.length} records`,
        "success"
      );

      this.displayDataPreview();
      this.hideProgress();
    } catch (error) {
      this.showStatus(`Error parsing file: ${error}`, "error");
      this.hideProgress();
    }
  }

  private parseExcelData(jsonData: any[]): ExcelRow[] {
    const parsedData: ExcelRow[] = [];

    if (jsonData.length < 2) {
      throw new Error(
        "Excel file must contain at least a header row and one data row"
      );
    }

    // Skip header row (assuming first row contains headers)
    for (let i = 1; i < jsonData.length; i++) {
      const row = jsonData[i];

      if (row && row.length >= 3) {
        parsedData.push({
          ID: row[0]?.toString() || "",
          Response: row[2]?.toString() || "",
          Notes: row[3]?.toString() || "",
        });
      }
    }

    return parsedData;
  }

  private displayDataPreview(): void {
    const previewSection = this._container.querySelector(
      ".preview-section"
    ) as HTMLDivElement;
    const dataPreview = this._container.querySelector(
      ".data-preview"
    ) as HTMLDivElement;

    let previewHTML = `
            <table class="data-table">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Response</th>
                        <th>Notes</th>
                    </tr>
                </thead>
                <tbody>
        `;

    // Show first 5 rows as preview
    const previewRows = this._parsedData.slice(0, 5);
    previewRows.forEach((row) => {
      previewHTML += `
                <tr>
                    <td>${this.escapeHtml(row.ID)}</td>
                    <td>${this.escapeHtml(row.Response)}</td>
                    <td>${this.escapeHtml(row.Notes)}</td>
                </tr>
            `;
    });

    previewHTML += `
                </tbody>
            </table>
        `;

    if (this._parsedData.length > 5) {
      previewHTML += `<p>... and ${this._parsedData.length - 5} more rows</p>`;
    }

    dataPreview.innerHTML = previewHTML;
    previewSection.style.display = "block";
  }

  private async processDataverseUpdates(): Promise<void> {
    if (this._parsedData.length === 0) {
      this.showStatus("No data to process", "error");
      return;
    }

    this.showStatus("Updating Dataverse records...", "info");
    this.showProgress(0);

    const results: UpdateResult[] = [];
    const totalRecords = this._parsedData.length;

    for (let i = 0; i < this._parsedData.length; i++) {
      const row = this._parsedData[i];

      try {
        const result = await this.updateDataverseRecord(row);
        results.push(result);

        const progress = Math.round(((i + 1) / totalRecords) * 100);
        this.showProgress(progress);
      } catch (error) {
        results.push({
          success: false,
          recordId: row.ID,
          error: error?.toString() || "Unknown error",
        });
      }
    }

    this.hideProgress();
    this.displayResults(results);
  }

  private async updateDataverseRecord(row: ExcelRow): Promise<UpdateResult> {
    try {
      // First, find the existing record by afsdc_name
      const query = `?$filter=afsdc_name eq '${row.ID}'`;
      const retrieveResponse =
        await this._context.webAPI.retrieveMultipleRecords(
          "afsdc_questionresponseinstance",
          query
        );

      if (retrieveResponse.entities.length === 0) {
        return {
          success: false,
          recordId: row.ID,
          error: "Record not found",
        };
      }

      const existingRecord = retrieveResponse.entities[0];

      // Update the record
      const updateData = {
        afsdc_response: row.Response,
        afsdc_comments: row.Notes,
      };

      await this._context.webAPI.updateRecord(
        "afsdc_questionresponseinstance",
        existingRecord.afsdc_questionresponseinstanceid,
        updateData
      );

      return {
        success: true,
        recordId: row.ID,
      };
    } catch (error) {
      return {
        success: false,
        recordId: row.ID,
        error: error?.toString() || "Update failed",
      };
    }
  }

  private displayResults(results: UpdateResult[]): void {
    const resultsSection = this._container.querySelector(
      ".results-section"
    ) as HTMLDivElement;

    const successCount = results.filter((r) => r.success).length;
    const failureCount = results.filter((r) => !r.success).length;

    let resultsHTML = `
            <div class="results-summary">
                <p><strong>Update Summary:</strong></p>
                <p>✅ Successful: ${successCount}</p>
                <p>❌ Failed: ${failureCount}</p>
            </div>
        `;

    if (failureCount > 0) {
      resultsHTML += `
                <div class="failed-records">
                    <h4>Failed Records:</h4>
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>Error</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

      results
        .filter((r) => !r.success)
        .forEach((result) => {
          resultsHTML += `
                    <tr>
                        <td>${this.escapeHtml(result.recordId)}</td>
                        <td>${this.escapeHtml(
                          result.error || "Unknown error"
                        )}</td>
                    </tr>
                `;
        });

      resultsHTML += `
                        </tbody>
                    </table>
                </div>
            `;
    }

    this._resultsDiv.innerHTML = resultsHTML;
    resultsSection.style.display = "block";

    if (successCount > 0) {
      this.showStatus(
        `Successfully updated ${successCount} records`,
        "success"
      );
    } else {
      this.showStatus("No records were updated", "error");
    }
  }

  private showStatus(
    message: string,
    type: "info" | "success" | "error"
  ): void {
    this._statusDiv.innerHTML = `<span class="status-${type}">${message}</span>`;
  }

  private showProgress(percentage: number): void {
    const progressFill = this._container.querySelector(
      ".progress-fill"
    ) as HTMLDivElement;
    const progressText = this._container.querySelector(
      ".progress-text"
    ) as HTMLSpanElement;

    this._progressDiv.style.display = "block";
    progressFill.style.width = `${percentage}%`;
    progressText.textContent = `${percentage}%`;
  }

  private hideProgress(): void {
    this._progressDiv.style.display = "none";
  }

  private escapeHtml(text: string): string {
    const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this._context = context;
    this._value = context.parameters.sampleProperty.raw || "";
  }

  public getOutputs(): IOutputs {
    return {
      sampleProperty: this._value,
    };
  }

  public destroy(): void {
    // Clean up resources
  }
}
