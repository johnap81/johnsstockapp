async function fetchStock() {
    const ticker = document.getElementById("tickerInput").value.toUpperCase();
    const resultDiv = document.getElementById("result");

    if (!ticker) {
        resultDiv.innerHTML = "Please enter a stock symbol.";
        return;
    }

    resultDiv.innerHTML = "Fetching price...";

    // Currency detection based on ticker suffix
    const currencyMap = {
        "NS": "INR",   // NSE India
        "BSE": "INR",  // BSE India
        "DE": "EUR",   // XETRA Germany
        "F": "EUR",    // Frankfurt
        "L": "GBP",    // London Stock Exchange
        "SW": "CHF",   // SIX Swiss Exchange
        "PA": "EUR",   // Euronext Paris
        "AS": "EUR",   // Euronext Amsterdam
        "BR": "EUR",   // Euronext Brussels
        "HE": "EUR",   // Helsinki
        "ST": "SEK",   // Stockholm
        "CO": "DKK"    // Copenhagen
    };

    // Default currency
    let currency = "USD";

    // Detect suffix
    if (ticker.includes(".")) {
        const suffix = ticker.split(".")[1];
        if (currencyMap[suffix]) {
            currency = currencyMap[suffix];
        }
    }

    try {
        const apiKey = "U7O1FRNUI7H9WWEA";  // your key
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
