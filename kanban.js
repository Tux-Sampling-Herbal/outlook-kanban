Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        loadTasks();
    }
});

function loadTasks() {
    Office.context.roamingSettings.getAsync("kanban_tasks", (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value) {
            const tasks = JSON.parse(asyncResult.value);
            renderBoard(tasks);
        } else {
            // Handle case where there are no tasks yet
            console.log("No tasks found or error loading settings.");
        }
    });
}

function renderBoard(tasks) {
    // Clear existing tasks
    document.getElementById("tasks-todo").innerHTML = "";
    document.getElementById("tasks-inprogress").innerHTML = "";
    document.getElementById("tasks-done").innerHTML = "";

    tasks.forEach(task => {
        const taskCard = document.createElement("div");
        taskCard.className = "task-card";
        taskCard.id = task.id;
        taskCard.innerHTML = `
            <h4>${task.title}</h4>
            <p><strong>Due:</strong> ${task.dueDate || 'N/A'}</p>
            <p><strong>To:</strong> ${task.assignedTo || 'N/A'}</p>
        `;

        if (task.status === "To Do") {
            document.getElementById("tasks-todo").appendChild(taskCard);
        } else if (task.status === "In Progress") {
            document.getElementById("tasks-inprogress").appendChild(taskCard);
        } else {
            document.getElementById("tasks-done").appendChild(taskCard);
        }
    });
}