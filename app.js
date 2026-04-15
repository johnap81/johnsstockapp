const apiKey = "U7O1FRNUI7H9WWEA";  // Alpha Vantage only for search
const workerURL = "https://red-grass-eae5.johnap81.workers.dev/?symbol=";

// SEARCH FUNCTION (Alpha Vantage)
async function searchStocks() {
    const input = document.getElementById("tickerInput");
    const query = input.value.trim();
    const suggestionsDiv = document.getElementById("suggestions");
    const resultDiv = document.getElementById("result");

    suggestionsDiv.innerHTML = "";
    resultDiv.innerHTML = "";

    if (!query) {
        resultDiv.innerHTML = "Please enter a stock name or symbol.";
        return;
    }

    resultDiv.innerHTML = "Searching...";

    try {
        const url = `https://www.alphavantage.co/query?function=SYMBOL_SEARCH&keywords=${encodeURIComponent(query)}&apikey=${apiKey}`;
        const response = await fetch(url);
        const data = await response.json();

        const matches = data["bestMatches"];

        if (!matches || matches.length === 0) {
            resultDiv.innerHTML = "No matches found.";
            return;
        }

        resultDiv.innerHTML = "Select a stock:";

        let html = "";
        matches.forEach(match => {
            const symbol = match["1. symbol"];
            const name = match["2. name"];
            const region = match["4. region"];
            const currency = match["8. currency"];

            html += `
                <button 
                    style="display:block;width:100%;margin:5px 0;padding:8px;border-radius:5px;border:1px solid #ccc;text-align:left;"
                    onclick="selectStock('${symbol}')">
                    <strong>${symbol}</strong> — ${name}<br>
                    <small>${region} • ${currency}</small>
                </button>
            `;
        });

        suggestionsDiv.innerHTML = html;
    } catch (error) {
        resultDiv.innerHTML = "Error searching stocks.";
    }
}

// WHEN USER CLICKS A STOCK
function selectStock(symbol) {
    const input = document.getElementById("tickerInput");
    const suggestionsDiv = document.getElementById("suggestions");

    input.value = symbol;
    suggestionsDiv.innerHTML = "";
    fetchStock();
}

// FETCH REAL PRICE FROM YOUR WORKER
async function fetchStock() {
    const ticker = document.getElementById("tickerInput").value.trim();
    const resultDiv = document.getElementById("result");

    if (!ticker) {
        resultDiv.innerHTML = "Please enter a stock symbol.";
        return;
    }

    resultDiv.innerHTML = "Fetching real price...";

    try {
        const response = await fetch(workerURL + ticker);
        const data = await response.json();

        if (data.error) {
            resultDiv.innerHTML = "Invalid stock symbol.";
            return;
        }

        resultDiv.innerHTML = `
            <strong>${data.symbol}</strong><br>
            ${data.name}<br><br>
            <strong>Price:</strong> ${data.price} ${data.currency}<br>
            <strong>Exchange:</strong> ${data.exchange}
        `;
    } catch (error) {
        resultDiv.innerHTML = "Error fetching data.";
    }
}
