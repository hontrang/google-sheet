function createTriggers() {
    const functionName = "layChiSoVnIndex"; // Tên của hàm bạn muốn chạy

    // Tạo trigger chạy theo thời gian
    const timeTrigger = ScriptApp.newTrigger(functionName)
        .timeBased()
        .everyMinutes(1)
        .create();

    Logger.log("Các trigger đã được tạo!");
}