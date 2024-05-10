    const conversation = document.getElementById('conversation');
    const userInput = document.getElementById('userInput');

    function postMessage(message, isUser) {
        const messageElem = document.createElement('div');
        messageElem.classList.add('message', isUser ? 'user' : 'ai');
        messageElem.textContent = message;
        conversation.appendChild(messageElem);
        conversation.scrollTop = conversation.scrollHeight; // Scroll to the bottom

        // Optional: Move focus to input after posting message
        userInput.focus();
    }

    userInput.addEventListener('keydown', function(event) {
        if (event.key === 'Enter') {
            event.preventDefault();
            const message = userInput.value.trim();
            if (message) {
                postMessage(userInput.value, true); // Display user message
                sendToAI(message); // Send user message to AI
                userInput.value = ''; // Clear input field
            }
        }
    });

    function sendToAI(userMessage) {
        // This function should send the user's message to your AI model
        // and then use postMessage to display the AI's response.

        const xhr = new XMLHttpRequest();
        xhr.open('POST', '/ai', true);
        xhr.setRequestHeader('Content-Type', 'application/json;charset=UTF-8');
        xhr.onreadystatechange = function() {
            if (xhr.readyState === XMLHttpRequest.DONE) {
                if (xhr.status === 200) {
                    const aiResponse = xhr.responseText;
                    postMessage(aiResponse, false); // Display AI response
                } else {
                    postMessage('Error: Unable to reach AI model.', false);
                }
            }
        };
        xhr.send(JSON.stringify({ message: userMessage }));
    }