const conversation = document.getElementById('conversation');
const userInput = document.getElementById('userInput');

function postMessage(message, isUser) {
    const messageElem = document.createElement('div');
    messageElem.classList.add('message', isUser ? 'user' : 'ai');
    messageElem.textContent = message;
    conversation.appendChild(messageElem);
    conversation.scrollTop = conversation.scrollHeight; // Scroll to the bottom
    userInput.focus();
}

userInput.addEventListener('keydown', function(event) {
    if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        const message = userInput.value.trim();
        if (message) {
            postMessage(message, true); // Display user message
            sendToAI(message); // Send user message to AI
            userInput.value = ''; // Clear input field
        }
    }
});

function sendToAI(userMessage) {
    fetch('/ai', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            prompt: userMessage
        })
    })
    .then(response => {
        if (!response.ok) {
            throw new Error('Network response was not ok');
        }
        return response.json();
    })
    .then(data => {
        if (data.error) {
            postMessage(data.error, false);
        } else {
            postMessage(data.response, false);
        }
    })
    .catch(error => {
        console.error('Error:', error);
        postMessage('与服务器通信时发生错误：' + error.message, false);
    });
}