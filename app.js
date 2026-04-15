const apiKey = "U7O1FRNUI7H9WWEA";  // your Alpha Vantage key

function detectCurrency(ticker) {
    const currencyMap = {
        "NS": "INR",   // NSE India
        "BSE": "INR",  // BSE India
        "DE": "EUR",   // XETRA Germany
        "F": "EUR",    // Frankfurt
        "L": "GBP",    // London
        "SW": "CHF",   // SIX Swiss
        "PA": "EUR",   // Paris
        "AS": "EUR",   // Amsterdam
        "BR": "EUR",   // Brussels
        "HE": "EUR",   // Helsinki
        "ST": "SEK",   // Stockholm
        "CO": "DKK"    // Copenhagen
    };

    let currency = "USD";

    if (ticker.includes(".")) {
        const suffix = ticker.split(".")[1];
        if (currencyMap[suffix]) {
            currency = currencyMap[suffix];
        }
    }

    return currency;
}

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

function selectStock(symbol) {
    const input = document.getElementById("tickerInput");
    const suggestionsDiv = document.getElementById("suggestions");

    input.value = symbol;
    suggestionsDiv.innerHTML = "";
    fetchStock();
}

async function fetchStock() {
    const ticker = document.getElementById("tickerInput").value.toUpperCase();
    const resultDiv = document.getElementById("result");

    if (!ticker) {
        resultDiv.innerHTML = "Please enter a stock symbol.";
        return;
    }

    resultDiv.innerHTML = "Fetching price...";

    const currency = detectCurrency(ticker);

    try {
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
            Price: ${price} ${currency}<br>
            Change: ${change} (${percent})
        `;
    } catch (error) {
        resultDiv.innerHTML = "Error fetching data.";
    }
}
