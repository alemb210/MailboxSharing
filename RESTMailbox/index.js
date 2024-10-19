document.getElementById('mailForm').addEventListener('submit', function (event) {
    event.preventDefault();
    const mailbox = document.getElementById('mailbox').value;
    const admin = document.getElementById('admin').value;
    const days = parseInt(document.getElementById('days').value, 10); // Ensure days is an integer
    const functionAppURL = "https://functionAppName.azurewebsites.net/api/FunctionName?code=FunctionKey"; //Replace with your function app URL from Azure
    fetch(functionAppURL, { //send POST request with information from forms
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ mailbox, admin, days }) 
    })
        .then(response => response.json())
        .then(data => console.log(data))
        .catch(error => console.error('Error:', error));
});


