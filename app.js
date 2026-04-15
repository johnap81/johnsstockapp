async function fetchStock() {
    const ticker = document.getElementById("tickerInput").value.toUpperCase();
    const resultDiv = document.getElementById("result");

    if (!ticker) {
        resultDiv.innerHTML = "Please enter a stock symbol.";
        return;
    }

    resultDiv.innerHTML = "Fetching price...";

    try {
        const apiKey = "YOUR_API_KEY_HERE";  // ← Replace this with your real key

        const url = `https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=${ticker}&apikey=${apiKey}`;

        const response = await fetch(url);
        const data = await response.json();

        if (!data["Global Quote"] || !data["Global Quote"]["05. price"]) {
            resultDiv.innerHTML = "Invalid stock symbol or API limit reached.";
            return;
        }

        const price = data["Global Quote"]["05. price"];
        const change = data["Global Quote"]["09. change"];
        const percent = data["Global Quote"]["10. change percent"];

        resultDiv.innerHTML = `
            <strong>${ticker}</strong><br>
            Price: ${price} USD<br>
            Change: ${change} (${percent})
        `;
    } catch (error) {
        resultDiv.innerHTML = "Error fetching data.";
    }
}
