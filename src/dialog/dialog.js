/* globals Office */
(function () {
  function bytesToMb(b) { return (b / (1024 * 1024)); }

  Office.initialize = () => {
    const params = new URLSearchParams(window.location.search);
    const thresholdMb = Number(params.get("thresholdMb") || 0);
    const rawBytes = Number(params.get("rawBytes") || 0);
    const adjBytes = Number(params.get("adjustedBytes") || 0);

    document.getElementById("threshold").textContent = thresholdMb.toFixed(0);
    document.getElementById("rawMb").textContent = bytesToMb(rawBytes).toFixed(2) + " MB";
    document.getElementById("adjMb").textContent = bytesToMb(adjBytes).toFixed(2) + " MB";

    document.getElementById("cancel").addEventListener("click", () => {
      Office.context.ui.messageParent("cancel");
    });
    document.getElementById("continue").addEventListener("click", () => {
      Office.context.ui.messageParent("continue");
    });
  };
})();
