document.getElementById("sendBtn").addEventListener("click", async function () {
    var userCommand = document.getElementById("messageInput").value;
    document.getElementById("messageInput").value = "";

    var chatBox = document.getElementById("chatBox");
    var newMessage = document.createElement("p");
    newMessage.textContent = "You: " + userCommand;
    chatBox.appendChild(newMessage);

    // Send the user command to the Python backend
    const response = await fetch("http://127.0.0.1:5000/run-command", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
        },
        body: JSON.stringify({ command: userCommand }),
    });

    const result = await response.json();

    var botMessage = document.createElement("p");
    botMessage.textContent = "Bot: " + result.message;
    chatBox.appendChild(botMessage);
    chatBox.scrollTop = chatBox.scrollHeight;

    if (result.js_code) {
    await Excel.run(async (context) => {
        try {
            eval(result.js_code); // Evaluate the received JavaScript code
            await context.sync(); // Ensure changes are applied
        } catch (error) {
            console.log("Eval error:", error);  // Improved error reporting
            var errorMessage = document.createElement("p");
            errorMessage.textContent = "Error executing JS code: " + error.message;
            chatBox.appendChild(errorMessage);
            chatBox.scrollTop = chatBox.scrollHeight;
        }
    });
}

});
