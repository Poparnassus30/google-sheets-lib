/**
 * Exemple: récupérer le prix TAO (spot) via CoinGecko (simple)
 * NOTE: API publique, pratique pour tester.
 */

function getTaoPriceEUR() {
  const url = "https://api.coingecko.com/api/v3/simple/price?ids=bittensor&vs_currencies=eur";
  const data = fetchJson(url);
  return data?.bittensor?.eur ?? null;
}

function getTaoPriceUSD() {
  const url = "https://api.coingecko.com/api/v3/simple/price?ids=bittensor&vs_currencies=usd";
  const data = fetchJson(url);
  return data?.bittensor?.usd ?? null;
}
