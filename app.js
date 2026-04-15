async function fetchStock() {
    const ticker = document.getElementById("tickerInput").value.toUpperCase();
    const resultDiv = document.getElementById("result");

    if (!ticker) {
        resultDiv.innerHTML = "Please enter a stock symbol.";
        return;
    }

    resultDiv.innerHTML = "Fetching price...";

    try {
        const response = await fetch(`https://query1.finance.yahoo.com/v7/finance/quote?symbols=${ticker}`);
        const data = await response.json();

        if (!data.quoteResponse.result.length) {
            resultDiv.innerHTML = "Invalid stock symbol.";
            return;
        }

        const stock = data.quoteResponse.result[0];
        const price = stock.regularMarketPrice;
        const currency = stock.currency;

        resultDiv.innerHTML = `
            <strong>${ticker}</strong><br>
            Price: ${price} ${currency}
        `;
    } catch (error) {
        resultDiv.innerHTML = "Error fetching data.";
    }
}
