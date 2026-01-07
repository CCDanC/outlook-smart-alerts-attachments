
/* globals Office */
(function () {
  const THRESHOLD_MB = 10;         // <-- set your policy value here
  const OVERHEAD_FACTOR = 1.33;    // transport overhead
  const EXCLUDE_CLOUD_ATTACHMENTS = true;
  const DIALOG_URL = "https://ccdanc.github.io/outlook-smart-alerts-attachments/src/dialog/dialog.html";
  const DIALOG_HEIGHT = 45; // percent
  const DIALOG_WIDTH = 40;  // percent

  Office.initialize = () => {};

  async function onMessageSendHandler(event) {
    try {
      const item = Office.context.mailbox.item;
      const attachments = await getAttachmentsAsync(item);
      const totalBytes = sumAttachmentBytes(attachments, EXCLUDE_CLOUD_ATTACHMENTS);
      const adjustedBytes = Math.round(totalBytes * OVERHEAD_FACTOR);
      const thresholdBytes = mbToBytes(THRESHOLD_MB);

      if (adjustedBytes > thresholdBytes) {
        const decision = await showDecisionDialog({
          thresholdMb: THRESHOLD_MB,
          rawBytes: totalBytes,
          adjustedBytes
        });
        event.completed({ allowEvent: decision !== "cancel" });
        return;
      }
      event.completed({ allowEvent: true });
    } catch (e) {
      console.error("Smart Alerts error:", e);
      event.completed({ allowEvent: true });
    }
  }

  window.onMessageSendHandler = onMessageSendHandler;

  function mbToBytes(mb) { return Math.round(mb * 1024 * 1024); }
  function bytesToMb(b) { return (b / (1024 * 1024)); }

  function sumAttachmentBytes(attachments, excludeCloud) {
    let total = 0;
    (attachments || []).forEach(att => {
      const isCloud = att.attachmentType === Office.AttachmentType.Cloud;
      if (excludeCloud && isCloud) return;
      const size = typeof att.size === "number" ? att.size : 0;
      if (att.attachmentType === Office.AttachmentType.File ||
          att.attachmentType === Office.AttachmentType.Item) {
        total += size;
      }
    });
    return total;
  }

  function getAttachmentsAsync(item) {
    return new Promise((resolve, reject) => {
      item.getAttachmentsAsync(result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || []);
        } else {
          reject(result.error);
        }
      });
    });
  }

  function showDecisionDialog(payload) {
    return new Promise((resolve) => {
      const query = new URLSearchParams({
        thresholdMb: String(payload.thresholdMb),
        rawBytes: String(payload.rawBytes),
        adjustedBytes: String(payload.adjustedBytes)
      }).toString();

      Office.context.ui.displayDialogAsync(
        `${DIALOG_URL}?${query}`,
        { height: DIALOG_HEIGHT, width: DIALOG_WIDTH, displayInIframe: true },
        (asyncResult) => {
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            resolve("continue"); return;
          }
          const dialog = asyncResult.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            const decision = (arg && arg.message || "").toLowerCase();
            dialog.close();
            resolve(decision === "cancel" ? "cancel" : "continue");
          });
          dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
            dialog.close();
            resolve("cancel");
          });
        }
      );
    });
  }
})();
