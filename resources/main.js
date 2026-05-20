"use strict";

(() => {
  class LeaveAppRecord {
    #applicationIdAddress;

    #applicationId;

    #status;

    #employeeId;

    #employeeName;

    #leaveStartTime;

    #leaveEndTime;

    constructor({
      applicationIdAddress,
      applicationId,
      status,
      employeeId,
      employeeName,
      leaveStartTime,
      leaveEndTime,
    }) {
      this.#applicationIdAddress = applicationIdAddress;
      this.#applicationId = applicationId;
      this.#status = status;
      this.#employeeId = employeeId;
      this.#employeeName = employeeName;
      this.#leaveStartTime = leaveStartTime;
      this.#leaveEndTime = leaveEndTime;
    }

    static fromRow(row) {
      if (!row) return null;

      const applicationIdAddress = "foo";
      const applicationId = row[1].v.trim();
      const status = row[4].v.trim();
      const employeeId = row[7].v.trim();
      const employeeName = row[8].v.trim();

      const leaveTime = row[12].v.trim();
      const { startTime, endTime } =
        LeaveAppRecord.#parseLeaveTimeRange(leaveTime);

      return applicationIdAddress &&
        applicationId &&
        status &&
        employeeId &&
        employeeName &&
        startTime &&
        endTime
        ? new LeaveAppRecord({
            applicationIdAddress,
            applicationId,
            status,
            employeeId,
            employeeName,
            leaveStartTime: startTime,
            leaveEndTime: endTime,
          })
        : null;
    }

    static #parseLeaveTimeRange(leaveTime) {
      const result = leaveTime.match(/\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2}/g);

      if (!result || result.length != 2) {
        console.warn(`Invalid leave time range: ${leaveTime}`);

        return {
          startTime: null,
          endTime: null,
        };
      }

      const startTime = parseDateTime(result[0]);
      const endTime = parseDateTime(result[1]);

      if (!startTime) console.warn(`Invalid start time format: ${result[0]}`);

      if (!endTime) console.warn(`Invalid end time format: ${result[1]}`);

      return {
        startTime,
        endTime,
      };
    }

    get applicationIdAddress() {
      return this.#applicationIdAddress;
    }

    get applicationId() {
      return this.#applicationId;
    }

    get status() {
      return this.#status;
    }

    get employeeId() {
      return this.#employeeId;
    }

    get employeeName() {
      return this.#employeeName;
    }

    get leaveStartTime() {
      return this.#leaveStartTime;
    }

    get leaveEndTime() {
      return this.#leaveEndTime;
    }
  }

  dayjs.extend(dayjs_plugin_customParseFormat);
  dayjs.extend(dayjs_plugin_isBetween);

  const workForm = document.getElementById("workForm");

  const excelInput = workForm.elements.excel;
  excelInput.addEventListener("change", (e) => {
    workForm.elements.check.disabled = e.currentTarget.files.length == 0;
    resetDownloadButton(workForm);
  });

  const checkButton = workForm.elements.check;
  checkButton.disabled = workForm.elements.excel.files.length == 0;
  checkButton.addEventListener("click", (e) => {
    const spinner = e.currentTarget.querySelector("[data-spinner]");
    spinner.classList.remove("d-none");

    resetDownloadButton(workForm);

    const excelInput = workForm.elements.excel;
    excelInput.classList.remove("is-invalid");

    loadWorkbook(excelInput)
      .then((wb) => {
        process(wb);

        const url = createDownloadUrl(wb);
        downloadButton.dataset.url = url;
        downloadButton.disabled = !url;
      })
      .catch((e) => {
        console.error(e);
        excelInput.classList.add("is-invalid");
      })
      .finally(() => spinner.classList.add("d-none"));
  });

  const downloadButton = workForm.elements.download;
  downloadButton.disabled = !downloadButton.dataset.url;
  downloadButton.addEventListener("click", downloadWorkbook);

  function resetDownloadButton(workForm) {
    const downloadButton = workForm.elements.download;
    downloadButton.disabled = true;

    if (downloadButton.dataset.url) {
      URL.revokeObjectURL(downloadButton.dataset.url);
      downloadButton.dataset.url = "";
    }
  }

  function loadWorkbook(excelInput) {
    if (excelInput.files.length == 0) throw new Error("No file is choosed");

    return excelInput.files[0]
      .arrayBuffer()
      .then((buffer) => XLSX.read(buffer, { dense: true }));
  }

  function createDownloadUrl(workbook) {
    const buffer = XLSX.writeXLSX(workbook, {
      type: "array",
      cellStyles: true,
    });

    return URL.createObjectURL(
      new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
    );
  }

  function downloadWorkbook(event) {
    const downloadButton = event.currentTarget;
    const url = downloadButton.dataset.url;

    if (!url) return;

    const download = document.createElement("a");
    download.download = "檢查結果.xlsx";
    download.href = url;
    download.click();
  }

  function process(workbook) {
    const leaveData = getLeaveApplicationRecords(workbook);
    processAbnormalPunchRecords(workbook, leaveData);
  }

  function getLeaveApplicationRecords(workbook) {
    const leaveSheet = workbook.Sheets["請假單"];

    if (!leaveSheet) throw new Error("請假單 work sheet not found");

    const leaveData = new Map();

    leaveSheet["!data"].forEach((row, index) => {
      if (index < 1) return;

      const record = LeaveAppRecord.fromRow(row);

      console.debug(record);

      const { valid, reason } = validateLeaveApplication(record);

      if (!valid) {
        XLSX.utils.sheet_add_aoa(leaveSheet, [[reason]], {
          origin: { r: index, c: 18 },
        });

        return;
      }

      const applications = leaveData.get(record.employeeId) || [];
      applications.push(record);

      leaveData.set(record.employeeId, applications);
    });

    return leaveData;
  }

  function validateLeaveApplication(record) {
    if (!record) return { valid: false, reason: "無效的資料" };

    const { status } = record;

    if (status.includes("抽單"))
      return { valid: false, reason: "跳過，流程狀態為抽單" };

    if (status.includes("駁回"))
      return { valid: false, reason: "跳過，流程狀態為駁回" };

    return { valid: true, reason: null };
  }

  function processAbnormalPunchRecords(workbook, leaveData) {
    const punchSheet = workbook.Sheets["打卡紀錄"];

    if (!punchSheet) throw new Error("打卡紀錄 work sheet not found");

    punchSheet["!data"].forEach((_, index) => {
      if (index < 6) return;

      const employeeId = getCellValue(punchSheet, index, 0).trim();
      const employeeName = getCellValue(punchSheet, index, 1).trim();
      const punchTime = getCellValue(punchSheet, index, 8).trim();
      const parsedPunchTime = parseDateTime(punchTime);
      let processedResult = "";

      console.debug(
        "employeeId: %s, employeeName: %s, punchTime: %s",
        employeeId,
        employeeName,
        punchTime,
      );

      if (!parsedPunchTime) {
        processedResult = "時間格式異常: " + punchTime;
        XLSX.utils.sheet_add_aoa(punchSheet, [[processedResult]], {
          origin: { r: index, c: 22 },
        });
        return;
      }

      const applications = leaveData.get(employeeId);

      if (!applications) {
        processedResult = "查無請假資料";
        XLSX.utils.sheet_add_aoa(punchSheet, [[processedResult]], {
          origin: { r: index, c: 22 },
        });
        return;
      }

      const app = applications.find((app) =>
        parsedPunchTime.isBetween(
          app.leaveStartTime,
          app.leaveEndTime,
          null,
          "[]",
        ),
      );

      if (app) {
        console.info(
          "%s has leave application(applicationId=%s, employeeId=%s, punchTime=%s, leaveStartTime=%s, leaveEndTime=%s",
          app.employeeName,
          app.applicationId,
          employeeId,
          parsedPunchTime.format(),
          app.leaveStartTime.format(),
          app.leaveEndTime.format(),
        );

        processedResult = `已請假，申請單編號為${app.applicationId}`;
        XLSX.utils.sheet_add_aoa(punchSheet, [[processedResult]], {
          origin: { r: index, c: 22 },
        });
      } else {
        processedResult = "查無請假資料";
        XLSX.utils.sheet_add_aoa(punchSheet, [[processedResult]], {
          origin: { r: index, c: 22 },
        });
      }
    });
  }

  function getCellValue(sheet, row, col) {
    const merges = sheet["!merges"];
    const rows = sheet["!data"];
    const value = rows[row]?.[col]?.v;

    if (!merges || merges.length == 0) return value;

    const merge = merges.find(
      ({ s, e }) => row >= s.r && row <= e.r && col >= s.c && col <= e.c,
    );

    return merge ? rows[merge.s.r]?.[merge.s.c]?.v : value;
  }

  function parseDateTime(dateTime) {
    if (!dateTime) return null;

    const result = dayjs(dateTime, "YYYY/MM/DD HH:mm", true);

    return result.isValid() ? result : null;
  }
})();
