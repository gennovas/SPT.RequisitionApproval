Office.initialize = function () {
    // Optional init logic
};

// Approve
function approveFunction(event) {
    handleReply("approved");
    event.completed();
}

// Reject Option 1
function rejectOption1Function(event) {
    handleReply("rejected", "ยอดไม่ถูกต้อง");
    event.completed();
}

// Reject Option 2
function rejectOption2Function(event) {
    handleReply("rejected", "ขั้นตอนไม่ถูกต้อง");
    event.completed();
}

// Reject Option 3 (custom reason)
function rejectOption3Function(event) {
    const reason = prompt("กรุณาใส่เหตุผลเพิ่มเติม:");
    handleReply("rejected", reason);
    event.completed();
}

// Common Reply Function
function handleReply(status, reasonCode) {
    try {
        const subject = Office.context.mailbox.item.subject || "";
        let body = "";

        if (status === "approved") {
            body = `${subject} was approved.`;
        } else if (status === "rejected") {
            body = `${subject} was rejected with reason code ${reasonCode}`;
        }

        // ใช้ displayReplyForm แทน replyAsync
        Office.context.mailbox.displayReplyForm({
            htmlBody: body
        });

    } catch (error) {
        console.error("Error in handleReply:", error);
    }
}