async function fetchStock() {
    const ticker = document.getElementById("tickerInput").value.toUpperCase();
    const resultDiv = document.getElementById("result");

    if (!ticker) {
        resultDiv.innerHTML = "Please enter a stock symbol.";
        return;
    }

    resultDiv.innerHTML = "Fetching price...";

    try {
        const response = await fetch(`https://financialmodelingprep.com/api/v3/quote/${ticker}?apikey=demo`);
        const data = await response.json();

        if (!data || data.length === 0) {
            resultDiv.innerHTML = "Invalid stock symbol.";
            return;
        }

        const stock = data[0];
        const price = stock.price;
        const change = stock.change;
        const percent = stock.changesPercentage;

        resultDiv.innerHTML = `
            <strong>${ticker}</strong><br>
            Price: ${price} USD<br>
            Change: ${change} (${percent}%)
        `;
    } catch (error) {
        resultDiv.innerHTML = "Error fetching data.";
    }
}
