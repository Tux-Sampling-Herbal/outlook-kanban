Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("saveTask").onclick = saveTask;
        populateTaskForm();
    }
});

function populateTaskForm() {
    const item = Office.context.mailbox.item;
    document.getElementById("taskTitle").value = item.subject;

    item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            document.getElementById("taskDescription").value = result.value.trim().substring(0, 500); // Limit description length
        }
    });
}

function saveTask() {
    const task = {
        id: "task_" + new Date().getTime(),
        title: document.getElementById("taskTitle").value,
        description: document.getElementById("taskDescription").value,
        dueDate: document.getElementById("dueDate").value,
        status: document.getElementById("status").value,
        assignedTo: document.getElementById("assignedTo").value,
        comments: document.getElementById("comments").value,
        sourceEmailId: Office.context.mailbox.item.itemId
    };

    // Retrieve existing tasks, add the new one, and save back
    Office.context.roamingSettings.getAsync("kanban_tasks", (asyncResult) => {
        let tasks = [];
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value) {
            tasks = JSON.parse(asyncResult.value);
        }

        tasks.push(task);

        Office.context.roamingSettings.setAsync("kanban_tasks", JSON.stringify(tasks), (setResult) => {
            if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                showNotification("Task saved successfully!");
            } else {
                showNotification("Error saving task.", true);
            }
        });
        Office.context.roamingSettings.saveAsync();
    });
}

function showNotification(message, isError = false) {
    const notification = document.getElementById("notification");
    notification.textContent = message;
    notification.className = "notification show " + (isError ? "error" : "success");
    setTimeout(() => {
        notification.className = "notification";
    }, 3000);
}