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
                    <button class="upload-btn" disabled>Upload & Process</button>
                </div>
                
                <div class="status-section">
                    <div class="status-message"></div>
                </div>
            </div>
            
            <!-- Modal -->
            <div class="modal-overlay" style="display: none;">
                <div class="modal-content">
                    <div class="modal-header">
                        <h3>Processing Excel File</h3>
                    </div>
                    <div class="modal-body">
                        <div class="modal-status"></div>
                        <div class="progress-bar">
                            <div class="progress-fill"></div>
                            <span class="progress-text">0%</span>
                        </div>
                        <div class="modal-results" style="display: none;"></div>
                    </div>
                    <div class="modal-footer" style="display: none;">
                        <button class="close-btn">Close & Refresh</button>
                    </div>
                </div>
            </div>
        `;

    this._fileInput = this._container.querySelector(
      ".file-input"
    ) as HTMLInputElement;
    this._uploadButton = this._container.querySelector(
      ".upload-btn"
    ) as HTMLButtonElement;
    this._statusDiv = this._container.querySelector(
      ".status-message"
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
      this.uploadAndProcessFile();
    });

    // Modal close button
    const closeBtn = this._container.querySelector(".close-btn");
    if (closeBtn) {
      closeBtn.addEventListener("click", () => {
        location.reload();
      });
    }
  }

  private async uploadAndProcessFile(): Promise<void> {
    const file = this._fileInput.files?.[0];
    if (!file) {
      this.showStatus("Please select a file first", "error");
      return;
    }

    // Show modal
    this.showModal();
    this.updateModalStatus("Parsing Excel file...", "info");
    this.updateModalProgress(0);

    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: "array" });

      console.log("Available sheets:", workbook.SheetNames);
      console.log("Workbook:", workbook);

      // Find the first visible sheet (ignore hidden sheets)
      let targetSheetName = workbook.SheetNames[0];

      // Look for a visible sheet if there are multiple
      for (const sheetName of workbook.SheetNames) {
        // Check if sheet is hidden in the workbook props
        const sheetIndex = workbook.SheetNames.indexOf(sheetName);
        const isHidden =
          workbook.Workbook?.Sheets?.[sheetIndex]?.Hidden !== undefined;

        if (!isHidden) {
          targetSheetName = sheetName;
          break;
        }
      }

      console.log("Using sheet:", targetSheetName);
      const worksheet = workbook.Sheets[targetSheetName];

      // Log sheet properties
      console.log(
        "Sheet properties:",
        Object.keys(worksheet).filter((key) => key.startsWith("!"))
      );

      // Convert to JSON with proper options
      const jsonData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: "",
        blankrows: false,
        raw: false, // This ensures values are converted to strings
      });

      console.log("Excel sheet data (first 10 rows):", jsonData.slice(0, 10));
      console.log("Total rows in Excel:", jsonData.length);

      this.updateModalProgress(30);

      // Parse the data
      this._parsedData = this.parseExcelData(jsonData);

      if (this._parsedData.length === 0) {
        throw new Error(
          "No valid data found in Excel file. Please check your file format."
        );
      }

      console.log(
        "Final parsed data (first 5 records):",
        this._parsedData.slice(0, 5)
      );

      this.updateModalProgress(50);
      this.updateModalStatus(
        `Found ${this._parsedData.length} records to process. Starting updates...`,
        "info"
      );

      // Process immediately without preview
      await this.processDataverseUpdates();
    } catch (error) {
      console.error("Error processing file:", error);
      this.updateModalStatus(`Error: ${error}`, "error");
      this.showModalFooter();
    }
  }

  private parseExcelData(jsonData: any[]): ExcelRow[] {
    const parsedData: ExcelRow[] = [];

    console.log("Raw Excel data:", jsonData);

    if (jsonData.length < 2) {
      throw new Error(
        "Excel file must contain at least a header row and one data row"
      );
    }

    // Skip header row (assuming first row contains headers)
    for (let i = 1; i < jsonData.length; i++) {
      const row = jsonData[i];
      console.log(`Row ${i}:`, row);

      if (row && Array.isArray(row) && row.length >= 1) {
        // Handle different possible column arrangements
        let id = "";
        let response = "";
        let notes = "";

        // Try to find ID (should be in first column)
        if (row[0] !== undefined && row[0] !== null && row[0] !== "") {
          id = row[0].toString().trim();
        }

        // Response could be in column 2 or 3 (index 1 or 2)
        if (
          row.length > 2 &&
          row[2] !== undefined &&
          row[2] !== null &&
          row[2] !== ""
        ) {
          response = row[2].toString().trim();
        } else if (
          row.length > 1 &&
          row[1] !== undefined &&
          row[1] !== null &&
          row[1] !== ""
        ) {
          response = row[1].toString().trim();
        }

        // Notes could be in column 3 or 4 (index 2 or 3)
        if (
          row.length > 3 &&
          row[3] !== undefined &&
          row[3] !== null &&
          row[3] !== ""
        ) {
          notes = row[3].toString().trim();
        } else if (
          row.length > 2 &&
          row[2] !== undefined &&
          row[2] !== null &&
          row[2] !== "" &&
          !response
        ) {
          notes = row[2].toString().trim();
        }

        // Only add if we have at least an ID
        if (id) {
          parsedData.push({
            ID: id,
            Response: response,
            Notes: notes,
          });
          console.log(
            `Added record: ID=${id}, Response=${response}, Notes=${notes}`
          );
        }
      }
    }

    console.log("Parsed data:", parsedData);
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
      this.updateModalStatus("No data to process", "error");
      this.showModalFooter();
      return;
    }

    const results: UpdateResult[] = [];
    const totalRecords = this._parsedData.length;

    for (let i = 0; i < this._parsedData.length; i++) {
      const row = this._parsedData[i];

      this.updateModalStatus(
        `Processing record ${i + 1} of ${totalRecords}: ${row.ID}`,
        "info"
      );

      try {
        const result = await this.updateDataverseRecord(row);
        results.push(result);

        const progress = Math.round(((i + 1) / totalRecords) * 100);
        this.updateModalProgress(progress);
      } catch (error) {
        console.error(`Error updating record ${row.ID}:`, error);
        results.push({
          success: false,
          recordId: row.ID,
          error: error?.toString() || "Unknown error",
        });
      }
    }

    this.displayModalResults(results);
  }

  private showModal(): void {
    const modal = this._container.querySelector(
      ".modal-overlay"
    ) as HTMLDivElement;
    modal.style.display = "flex";
  }

  private hideModal(): void {
    const modal = this._container.querySelector(
      ".modal-overlay"
    ) as HTMLDivElement;
    modal.style.display = "none";
  }

  private updateModalStatus(
    message: string,
    type: "info" | "success" | "error"
  ): void {
    const statusDiv = this._container.querySelector(
      ".modal-status"
    ) as HTMLDivElement;
    statusDiv.innerHTML = `<span class="status-${type}">${message}</span>`;
  }

  private updateModalProgress(percentage: number): void {
    const progressFill = this._container.querySelector(
      ".modal-content .progress-fill"
    ) as HTMLDivElement;
    const progressText = this._container.querySelector(
      ".modal-content .progress-text"
    ) as HTMLSpanElement;

    progressFill.style.width = `${percentage}%`;
    progressText.textContent = `${percentage}%`;
  }

  private showModalFooter(): void {
    const footer = this._container.querySelector(
      ".modal-footer"
    ) as HTMLDivElement;
    footer.style.display = "block";
  }

  private displayModalResults(results: UpdateResult[]): void {
    const resultsDiv = this._container.querySelector(
      ".modal-results"
    ) as HTMLDivElement;

    const successCount = results.filter((r) => r.success).length;
    const failureCount = results.filter((r) => !r.success).length;

    let resultsHTML = `
            <div class="results-summary">
                <p><strong>Update Complete!</strong></p>
                <p>✅ Successful: ${successCount}</p>
                <p>❌ Failed: ${failureCount}</p>
            </div>
        `;

    if (failureCount > 0) {
      resultsHTML += `
                <div class="failed-records">
                    <h4>Failed Records:</h4>
                    <div class="failed-list">
            `;

      results
        .filter((r) => !r.success)
        .forEach((result) => {
          resultsHTML += `
                    <div class="failed-item">
                        <strong>${this.escapeHtml(
                          result.recordId
                        )}</strong>: ${this.escapeHtml(
            result.error || "Unknown error"
          )}
                    </div>
                `;
        });

      resultsHTML += `
                    </div>
                </div>
            `;
    }

    resultsDiv.innerHTML = resultsHTML;
    resultsDiv.style.display = "block";

    if (successCount > 0) {
      this.updateModalStatus(
        `Successfully updated ${successCount} records!`,
        "success"
      );
    } else {
      this.updateModalStatus("No records were updated", "error");
    }

    this.showModalFooter();
  }

  private async updateDataverseRecord(row: ExcelRow): Promise<UpdateResult> {
    try {
      console.log(`\n=== Processing Record: ${row.ID} ===`);
      console.log(`Response: "${row.Response}"`);
      console.log(`Notes: "${row.Notes}"`);

      // Clean the ID value
      const cleanId = row.ID.toString().trim();

      // Try different query approaches
      let retrieveResponse;

      try {
        // Method 1: Simple filter
        const query1 = `?$filter=afsdc_name eq '${cleanId}'&$select=afsdc_questionresponseinstanceid,afsdc_name`;
        console.log(`Trying query 1: ${query1}`);
        retrieveResponse = await this._context.webAPI.retrieveMultipleRecords(
          "afsdc_questionresponseinstances",
          query1
        );
        console.log(`Query 1 result:`, retrieveResponse);
      } catch (error1) {
        console.log(`Query 1 failed:`, error1);

        try {
          // Method 2: Try with different entity name
          const query2 = `?$filter=afsdc_name eq '${cleanId}'&$select=afsdc_questionresponseinstanceid,afsdc_name`;
          console.log(`Trying query 2 with different entity name: ${query2}`);
          retrieveResponse = await this._context.webAPI.retrieveMultipleRecords(
            "afsdc_questionresponseinstance",
            query2
          );
          console.log(`Query 2 result:`, retrieveResponse);
        } catch (error2) {
          console.log(`Query 2 failed:`, error2);

          // Method 3: Try without filter to see if entity exists
          try {
            const query3 = `?$select=afsdc_questionresponseinstanceid,afsdc_name&$top=5`;
            console.log(`Trying query 3 to test entity access: ${query3}`);
            const testResponse =
              await this._context.webAPI.retrieveMultipleRecords(
                "afsdc_questionresponseinstances",
                query3
              );
            console.log(
              `Entity test successful. Sample records:`,
              testResponse
            );

            // If we can access the entity, the specific record just doesn't exist
            return {
              success: false,
              recordId: row.ID,
              error: `Record '${cleanId}' not found in entity`,
            };
          } catch (error3) {
            console.log(`Entity access test failed:`, error3);
            return {
              success: false,
              recordId: row.ID,
              error: `Entity access error: ${
                (error3 as any)?.message || "Unknown error"
              }`,
            };
          }
        }
      }

      if (!retrieveResponse || retrieveResponse.entities.length === 0) {
        console.log(`No record found for ID: ${cleanId}`);
        return {
          success: false,
          recordId: row.ID,
          error: `Record not found: '${cleanId}'`,
        };
      }

      const existingRecord = retrieveResponse.entities[0];
      console.log(`Found existing record:`, existingRecord);

      // Prepare update data
      const updateData: any = {};

      if (row.Response && row.Response.toString().trim() !== "") {
        updateData.afsdc_response = row.Response.toString().trim();
      }

      if (row.Notes && row.Notes.toString().trim() !== "") {
        updateData.afsdc_comments = row.Notes.toString().trim();
      }

      // Only update if we have data to update
      if (Object.keys(updateData).length === 0) {
        console.log(`No update data for record: ${cleanId}`);
        return {
          success: false,
          recordId: row.ID,
          error: "No data to update (Response and Notes are empty)",
        };
      }

      console.log(
        `Updating record ${existingRecord.afsdc_questionresponseinstanceid} with:`,
        updateData
      );

      // Try update
      await this._context.webAPI.updateRecord(
        "afsdc_questionresponseinstance",
        existingRecord.afsdc_questionresponseinstanceid,
        updateData
      );

      console.log(`✓ Successfully updated record: ${cleanId}`);

      return {
        success: true,
        recordId: row.ID,
      };
    } catch (error) {
      console.error(`✗ Error updating record ${row.ID}:`, error);
      return {
        success: false,
        recordId: row.ID,
        error: `Update failed: ${
          (error as any)?.message || error?.toString() || "Unknown error"
        }`,
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
    this._value = context.parameters.value.raw || "";
  }

  public getOutputs(): IOutputs {
    return {
      value: this._value,
    };
  }

  public destroy(): void {
    // Clean up resources
  }
}
