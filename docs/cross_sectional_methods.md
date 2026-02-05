# Cross-sectional analysis methods and momentum augmentations

This document summarizes cross-sectional analytics with academic and industry backing, plus common augmentations for momentum strategies. The citations below point to well-known papers and practitioner research.

## Cross-sectional factor methods (beyond momentum)

- **Value (cross-sectional)**: Rank stocks by valuation signals such as book-to-market or earnings yield and buy “cheap” vs short “expensive.” Canonical evidence: Fama & French (1992, 1993). Industry usage: factor investing and smart beta products.
  - Sources:
    - Fama, Eugene F. & French, Kenneth R. (1992). “The Cross‑Section of Expected Stock Returns.”
    - Fama, Eugene F. & French, Kenneth R. (1993). “Common risk factors in the returns on stocks and bonds.”

- **Size (cross-sectional)**: Small‑cap stocks outperform large‑cap stocks on average; rank by market cap. Canonical evidence: Fama & French (1992, 1993).
  - Sources:
    - Fama, Eugene F. & French, Kenneth R. (1992). “The Cross‑Section of Expected Stock Returns.”
    - Fama, Eugene F. & French, Kenneth R. (1993). “Common risk factors in the returns on stocks and bonds.”

- **Profitability / Quality (cross-sectional)**: Rank by operating profitability, return on equity, or similar quality metrics; go long high‑quality, short low‑quality. Canonical evidence: Novy‑Marx (2013). Industry usage: “quality” and “profitability” factors.
  - Sources:
    - Novy‑Marx, Robert (2013). “The Other Side of Value: The Gross Profitability Premium.”

- **Investment / Asset Growth (cross-sectional)**: Rank by asset growth or investment; low investment tends to outperform high investment. Canonical evidence: Fama & French (2015) and earlier asset‑growth literature.
  - Sources:
    - Fama, Eugene F. & French, Kenneth R. (2015). “A five‑factor asset pricing model.”

- **Low Risk / Low Volatility (cross-sectional)**: Rank by total or idiosyncratic volatility / beta; long low‑risk, short high‑risk. Canonical evidence: Ang et al. (2006, 2009); Baker, Bradley & Wurgler (2011).
  - Sources:
    - Ang, Andrew; Hodrick, Robert J.; Xing, Yuhang; Zhang, Xiaoyan (2006). “The Cross‑Section of Volatility and Expected Returns.”
    - Ang, Andrew; Hodrick, Robert J.; Xing, Yuhang; Zhang, Xiaoyan (2009). “High Idiosyncratic Volatility and Low Returns.”
    - Baker, Malcolm; Bradley, Brendan; Wurgler, Jeffrey (2011). “Benchmarks as Limits to Arbitrage: Understanding the Low‑Volatility Anomaly.”

- **Liquidity (cross-sectional)**: Rank by liquidity measures (e.g., Amihud illiquidity) and go long more liquid, short less liquid (or vice‑versa depending on expected premium). Canonical evidence: Amihud (2002), Pastor & Stambaugh (2003).
  - Sources:
    - Amihud, Yakov (2002). “Illiquidity and Stock Returns: Cross‑Section and Time‑Series Effects.”
    - Pastor, Lubos; Stambaugh, Robert F. (2003). “Liquidity Risk and Expected Stock Returns.”

- **Earnings Momentum / Revisions (cross-sectional)**: Rank by analyst earnings revisions or standardized unexpected earnings. Canonical evidence: Chan, Jegadeesh & Lakonishok (1996). Industry: earnings‑momentum factors.
  - Sources:
    - Chan, Louis K. C.; Jegadeesh, Narasimhan; Lakonishok, Josef (1996). “Momentum Strategies.”

- **Carry / Yield (cross-sectional)**: Rank by dividend yield or other carry-like measures. Often used as a cross‑sectional signal in equities and across asset classes; documented in factor literature and practitioner research.
  - Sources:
    - Asness, Cliff S.; Moskowitz, Tobias J.; Pedersen, Lasse H. (2013). “Value and Momentum Everywhere.”

## Momentum augmentations (techniques to improve base momentum)

- **Skip‑month / skip‑week**: Skip the most recent period (e.g., 1 week or 1 month) to avoid short‑term reversal effects. Canonical: Jegadeesh & Titman (1993); evidence of short‑term reversal in Lehmann (1990).
  - Sources:
    - Jegadeesh, Narasimhan; Titman, Sheridan (1993). “Returns to Buying Winners and Selling Losers: Implications for Stock Market Efficiency.”
    - Lehmann, Bruce N. (1990). “Fads, Martingales, and Market Efficiency.”

- **Volatility scaling / risk parity sizing**: Scale momentum positions inversely by volatility to equalize risk contribution across names (or use target volatility). Common in practitioner factor portfolios.
  - Sources:
    - Moskowitz, Tobias J.; Ooi, Yao Hua; Pedersen, Lasse H. (2012). “Time Series Momentum.”

- **Sector / industry neutrality**: Compute ranks within sectors and/or neutralize industry exposures to avoid sector bets dominating. Common in industry factor implementations (e.g., AQR, MSCI/Barra).
  - Sources:
    - Asness, Cliff S.; Frazzini, Andrea (2013). “The Devil in HML’s Details.”

- **Multi‑horizon momentum (blend)**: Combine multiple lookback windows (e.g., 3‑, 6‑, 12‑month returns) or blend cross‑sectional and time‑series momentum.
  - Sources:
    - Moskowitz, Tobias J.; Ooi, Yao Hua; Pedersen, Lasse H. (2012). “Time Series Momentum.”
    - Asness, Cliff S.; Moskowitz, Tobias J.; Pedersen, Lasse H. (2013). “Value and Momentum Everywhere.”

- **Residual (idiosyncratic) momentum**: Rank on residual returns after removing factor exposures (e.g., market, industry) to isolate stock‑specific momentum.
  - Sources:
    - Blitz, David; Hanauer, Matthias (2019). “Residual Momentum.”

- **Liquidity / tradability filters**: Require minimum volume/price/market cap to reduce transaction costs and improve implementability.
  - Sources:
    - Korajczyk, Robert A.; Sadka, Ronnie (2004). “Are Momentum Profits Robust to Trading Costs?”

## Practitioner / industry sources

- **AQR**: Published factor research on value, momentum, quality, and low‑risk.
  - Sources:
    - Asness, Cliff S.; Moskowitz, Tobias J.; Pedersen, Lasse H. (2013). “Value and Momentum Everywhere.”

- **Two Sigma**: Publishes research on factor investing, systematic equity signals, and momentum‑like effects (papers and blogs vary by year).

---

**Note:** The analysis screen can expose these as modular cross‑sectional strategies: value, size, quality/profitability, investment, low‑volatility, liquidity, earnings‑momentum, and carry/yield, plus momentum augmentations like skip‑periods, volatility scaling, sector neutrality, and residual momentum.
