"""Build the research proposal (Bab I–III) as a .docx in the DEB SV UGM format.

Usage: python build.py OUT.docx FIGDIR [pages.json]
pages.json maps heading/caption text to printed page numbers (second pass for the
table of contents). Without it, page numbers print as "00".
"""
import json
import sys
from pathlib import Path

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

from docxlib import Thesis, add_runs

out, figdir = Path(sys.argv[1]), Path(sys.argv[2])
pages = json.loads(Path(sys.argv[3]).read_text()) if len(sys.argv) > 3 else {}
logo = figdir / "ugm_logo.png"


def pg(key):
    return pages.get(key, "00")


T = Thesis()
doc = T.doc

TITLE_ID = ("PENGARUH EARNINGS SURPRISE TERHADAP POST-EARNINGS ANNOUNCEMENT DRIFT: "
            "BUKTI DARI SAHAM LQ45 DI BURSA EFEK INDONESIA")
TITLE_EN = ("THE EFFECT OF EARNINGS SURPRISE ON POST-EARNINGS ANNOUNCEMENT DRIFT: "
            "EVIDENCE FROM LQ45 STOCKS ON THE INDONESIA STOCK EXCHANGE")

# =============================================================== COVER (no number)
T.centered("PROPOSAL TUGAS AKHIR", bold=True, space_after=12)
T.centered(TITLE_ID, bold=True, space_after=8)
T.centered(TITLE_EN, bold=True, italic=True, space_after=24)
T.centered("Final Research Proposal (Chapters I–III)", bold=True)
T.centered("Disusun untuk memenuhi tugas akhir mata kuliah Business Research Methods",
           bold=True)
T.centered("Dosen Pengampu", bold=True)
T.centered("Paramita Her Astuti, S.E., M.Sc.", bold=True)
T.centered("Dr. Raeef Abdullah Altamimi, M.Acc.", bold=True, space_after=18)
if logo.exists():
    par = doc.add_paragraph()
    par.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.add_run().add_picture(str(logo), width=Cm(5.5), height=Cm(5.71))
else:
    T.centered("[Lambang UGM, 5,5 cm × 5,71 cm]", italic=True, space_before=60,
               space_after=60)
T.centered("Oleh:", bold=True, space_before=18)
T.centered("Azim Muzakki Hasnan", bold=True)
T.centered("24/539405/SV/24597", bold=True, space_after=24)
T.centered("PROGRAM STUDI PERBANKAN", bold=True)
T.centered("DEPARTEMEN EKONOMIKA DAN BISNIS", bold=True)
T.centered("SEKOLAH VOKASI", bold=True)
T.centered("UNIVERSITAS GADJAH MADA", bold=True)
T.centered("2026", bold=True)

# =============================================================== FRONT MATTER (i, ii…)
T.new_section("lowerRoman", 1, WD_ALIGN_PARAGRAPH.CENTER)
T.front_title("DAFTAR ISI", toc=False)
toc_entries = [
    (0, "DAFTAR TABEL"), (0, "DAFTAR GAMBAR"), (0, "DAFTAR ISTILAH"),
    (0, "BAB I PENDAHULUAN"),
    (1, "1.1 Research Background"), (1, "1.2 Problem Definition"),
    (1, "1.3 Research Questions and Research Objectives"),
    (1, "1.4 Scope and Significance of the Study"), (1, "1.5 Research Novelty"),
    (0, "BAB II LITERATURE REVIEW AND THEORETICAL FRAMEWORK"),
    (1, "2.1 Literature Synthesis"),
    (1, "2.2 Theoretical Framework and Conceptual Model"),
    (1, "2.3 Hypotheses Development"),
    (0, "BAB III RESEARCH METHODOLOGY"),
    (1, "3.1 Research Design"), (1, "3.2 Operational Definition and Measurement"),
    (1, "3.3 Population and Sampling Design"),
    (1, "3.4 Data Collection Methods and Bias Reduction"),
    (1, "3.5 Data Analysis Plan"), (1, "3.6 Research Ethics and Thesis Timeline"),
    (0, "DAFTAR PUSTAKA"),
]
for lvl, text in toc_entries:
    T.toc_line(text, pg(text), level=lvl, size=12 if lvl == 0 else 11.5)

TABLE_LIST = [
    ("1.1", "Business Symptoms versus the Core Management Problem"),
    ("1.2", "Preliminary Evidence on Delayed Price Responses"),
    ("1.3", "Research Questions and Matching Research Objectives"),
    ("1.4", "Positioning of This Study against Prior Studies"),
    ("2.1", "Summary of Prior Empirical Studies"),
    ("2.2", "Theoretical Predictions for Earnings Surprise and PEAD"),
    ("2.3", "Identification of Research Variables"),
    ("2.4", "Summary of Hypotheses"),
    ("3.1", "Research Design Elements"),
    ("3.2", "Operational Definitions of Variables"),
    ("3.3", "Operationalization Matrix"),
    ("3.4", "Sample Selection Procedure"),
    ("3.5", "Data Sources and Collection Protocol"),
    ("3.6", "Potential Biases and Mitigation Steps"),
    ("3.7", "Hypothesis Testing Techniques and Decision Rules"),
    ("3.8", "Planned Robustness Checks"),
    ("3.9", "Thesis Completion Timeline"),
]
FIG_LIST = [
    ("2.1", "Conceptual Model of the Study"),
    ("3.1", "Research Procedure"),
    ("3.2", "Thesis Completion Gantt Chart"),
]

T.page_break()
T.front_title("DAFTAR TABEL")
for n, t in TABLE_LIST:
    T.toc_line(f"Table {n}. {t}", pg(f"Table {n}"), size=11.5)
T.p("")
T.front_title("DAFTAR GAMBAR")
for n, t in FIG_LIST:
    T.toc_line(f"Figure {n}. {t}", pg(f"Figure {n}"), size=11.5)

T.page_break()
T.front_title("DAFTAR ISTILAH")
glossary = [
    ("AR", "Abnormal return: actual stock return minus the benchmark (IHSG) return"),
    ("CAR", "Cumulative abnormal return over a stated window of trading days"),
    ("D1 … D5", "Earnings-surprise quintile portfolios (D1 = most negative, D5 = most positive)"),
    ("EMH", "Efficient Market Hypothesis"),
    ("EPS", "Earnings per share"),
    ("IDX / BEI", "Indonesia Stock Exchange (Bursa Efek Indonesia)"),
    ("IFRS", "International Financial Reporting Standards"),
    ("IHSG", "Indeks Harga Saham Gabungan (Jakarta Composite Index)"),
    ("KSEI", "Kustodian Sentral Efek Indonesia (Indonesian Central Securities Depository)"),
    ("LQ45", "IDX index of 45 stocks with high liquidity and large market capitalization"),
    ("MSCI", "Global index provider whose market classification guides international funds"),
    ("OJK", "Otoritas Jasa Keuangan (Financial Services Authority)"),
    ("PEAD", "Post-earnings announcement drift"),
    ("PSAK", "Pernyataan Standar Akuntansi Keuangan (Indonesian accounting standards)"),
    ("SID", "Single Investor Identification (investor account count published by KSEI)"),
    ("SUE", "Standardized unexpected earnings (earnings surprise)"),
    ("SUER", "Quintile rank of SUE within a calendar quarter (1 to 5)"),
]
for k, v in glossary:
    par = doc.add_paragraph()
    pf = par.paragraph_format
    pf.line_spacing = 1.0
    pf.space_after = Pt(5)
    pf.left_indent = Cm(3.0)
    pf.first_line_indent = Cm(-3.0)
    pf.tab_stops.add_tab_stop(Cm(3.0))
    add_runs(par, f"{k}\t{v}")

# =============================================================== BODY (1, 2, 3…)
T.new_section("decimal", 1, WD_ALIGN_PARAGRAPH.RIGHT)

# ------------------------------------------------------------------------ BAB I
T.chapter("BAB I", "PENDAHULUAN")
T.h2("1.1 Research Background")
T.p("The efficient market hypothesis states that, in a market that is efficient in the "
    "semi-strong form, security prices fully and quickly reflect all publicly available "
    "information (Fama, 1970). The event-study method was developed precisely to examine "
    "how fast prices adjust to such information (Fama et al., 1969). Quarterly earnings "
    "announcements are the most frequent, scheduled, and value-relevant disclosures a listed "
    "firm makes, so they provide a natural setting for testing this prediction: if the market "
    "is efficient, the price reaction to an earnings surprise should be completed within a "
    "few days of the announcement.")
T.p("The evidence has repeatedly contradicted this prediction. Ball and Brown (1968) were the "
    "first to show that stock prices keep moving in the direction of the earnings surprise after "
    "the announcement. Bernard and Thomas (1989, 1990) confirmed that stocks with the most "
    "positive surprises continue to earn positive abnormal returns, and stocks with the most "
    "negative surprises continue to earn negative abnormal returns, for about 60 trading days. "
    "They attribute this post-earnings announcement drift (PEAD) to investors failing to "
    "recognize what current earnings imply for future earnings. Fama (1998, p. 286) calls PEAD "
    "the “granddaddy of underreaction events”, and a replication fifty years after the original "
    "study still finds that post-announcement returns can be predicted to a degree from the "
    "announcement (Ball & Brown, 2019).")
T.p("PEAD is often described as a global phenomenon, but its size differs widely across "
    "countries. Hung et al. (2015) show that the drift is not significant in several markets, "
    "including Belgium, Denmark, Finland, France, Greece, Spain, India and Brazil, during "
    "certain periods, and that it weakened in some countries after mandatory adoption of "
    "International Financial Reporting Standards (IFRS). Whether PEAD exists in a given market "
    "is therefore an empirical question rather than an assumption. In emerging Asia, Ho et al. "
    "(2024) document PEAD in Vietnam using 20,491 firm-quarters from 2009 to 2020: buying the "
    "stocks with the highest earnings surprises and selling those with the lowest earns abnormal "
    "returns of 3.0%, 3.7% and 6.5% over 3, 30 and 60 trading days. The drift remains "
    "5.0% over 60 days even among the largest firms. Ho et al. (2024) link this partly to "
    "Vietnam’s weaker accounting standards, and they note that Indonesia, Thailand and the "
    "Philippines have accounting systems closer to international standards or adopted corporate "
    "governance guidelines earlier. Whether drift also appears in such a market is left open.")
T.p("Indonesia is a relevant setting for answering that question. Its accounting system is "
    "closer to international standards than Vietnam’s (Ho et al., 2024), while its investor "
    "base has grown very quickly. The Indonesian Central Securities Depository (KSEI) recorded "
    "about 14.8 million capital-market investor accounts (SID) at the end of 2024 and 20.12 "
    "million by 19 December 2025, an increase of 35% in one year; about 8.5 million of these "
    "accounts hold equities (Irawati, 2025). Markets with many new, "
    "individual investors are where limited attention is most likely to slow the processing of "
    "earnings news (DellaVigna & Pollet, 2009; Hirshleifer et al., 2009). Preliminary Indonesian "
    "evidence points in this direction. Nainggolan (2025) studies 621 IDX firms and finds "
    "significant abnormal returns in the first week after earnings releases and drift lasting "
    "up to 30 days. Wijayana and Achjari (2020) report that Indonesian investors need time to "
    "recognize the value implications of information technology investment announcements in an "
    "environment of thin trading. "
    "Ramdhani (2021) finds that the components of the earnings surprise are related to PEAD on "
    "the IDX over 2017–2019.")
T.p("The question has become urgent in 2026, because the quality of Indonesia’s market "
    "information environment is now under international scrutiny. Moody’s changed the outlook "
    "on Indonesia’s Baa2 sovereign rating to negative in February 2026, citing reduced "
    "predictability in policymaking that “points to weakening governance” (Moody’s Ratings, "
    "2026, p. 1), and Fitch changed its outlook on Indonesia’s BBB rating to negative in March "
    "2026, citing increasing policy uncertainty (Fitch Ratings, 2026). In October 2026, MSCI proposed reclassifying the MSCI Indonesia Indexes from "
    "Emerging Market to Standalone Market status. MSCI reported that investors had raised "
    "concerns about “persistent opacity in shareholding structures and suspected coordinated "
    "trading behavior” (MSCI, 2026, p. 1), and that despite reforms, including disclosure of shareholders above "
    "1% and a roadmap to a 15% minimum free float, concerns remained about market "
    "dysfunction (MSCI, 2026). MSCI will decide by June 2027. All of these signals concern "
    "the same underlying issue: whether public information is reflected in Indonesian stock "
    "prices quickly and fairly. PEAD is a direct, measurable indicator of that issue, and Ho "
    "et al. (2024) show that the drift is larger where ownership is concentrated in insiders "
    "and information is more opaque. Evidence on how efficiently earnings news is priced in "
    "Indonesia’s leading stocks is therefore relevant not only to investors but to the "
    "credibility of the market as a whole.")
T.p("These Indonesian studies, however, use samples that span the whole exchange, where many "
    "stocks trade thinly. Chordia et al. (2009) show that PEAD is concentrated in illiquid stocks "
    "and that transaction costs absorb 70–100% of the paper profit of trading on it, while "
    "Mendenhall (2004) shows that the drift is larger where arbitrage is riskier. A drift found "
    "in a whole-market sample may therefore be impossible to trade and may say little about how "
    "efficiently the most important stocks are priced. This study focuses instead on the LQ45 "
    "index, which consists of the 45 IDX stocks with the highest liquidity and large market "
    "capitalization and whose constituents the IDX reviews periodically, with major reviews in "
    "February and August. LQ45 stocks are the most closely followed "
    "by analysts, institutional investors and the media, so they are the hardest place in "
    "Indonesia to find underreaction. If drift exists here, it is economically meaningful "
    "because it can be traded at low cost; if it does not, that result is itself informative "
    "about how quickly earnings news is priced in the liquid segment of the market.")
T.p("The study follows the hallmarks of scientific research (Sekaran & Bougie, 2016). It is "
    "*purposive* because it has a single, clearly stated aim: to test whether earnings "
    "surprises are followed by a drift among LQ45 stocks. It is *rigorous* because it applies "
    "the portfolio and regression tests established by Bernard and Thomas (1989) and applied "
    "by Ho et al. (2024), with event windows and hypotheses fixed in advance in this proposal. "
    "It is *objective* and *replicable* because every variable is computed from public IDX data "
    "using stated formulas, so another researcher can reproduce the results.")

T.h2("1.2 Problem Definition")
T.p("Sekaran and Bougie (2016) distinguish the symptoms that managers observe from the "
    "underlying problem that research must address. Table 1.1 separates the observable symptoms "
    "in the Indonesian market from the core problem faced by investors who trade LQ45 stocks.")
T.caption_table("1.1", "Business Symptoms versus the Core Management Problem")
T.table(
    ["Level", "Description"],
    [
        ["Symptom 1",
         "Indonesian studies observe abnormal returns that continue for days or weeks after "
         "corporate and earnings announcements (Nainggolan, 2025; Wijayana & Achjari, 2020)."],
        ["Symptom 2",
         "The investor base grew by 35% in 2025 alone, from about 14.8 million to 20.12 million "
         "accounts (Irawati, 2025); many new individual investors are prone to limited "
         "attention."],
        ["Symptom 3",
         "Earnings releases cluster close to regulatory reporting deadlines, so many reports "
         "reach the market on the same days, a condition linked to slower price responses "
         "(Hirshleifer et al., 2009)."],
        ["Symptom 4",
         "Existing Indonesian evidence comes from whole-market samples dominated by thinly "
         "traded stocks, so investors cannot tell whether the observed drift is tradable."],
        ["Symptom 5",
         "In 2026, Moody’s and Fitch moved Indonesia’s rating outlook to negative and MSCI "
         "proposed downgrading Indonesia to Standalone Market status, citing opaque "
         "shareholding structures and market dysfunction (Fitch Ratings, 2026; Moody’s Ratings, "
         "2026; MSCI, 2026)."],
        ["Core management problem",
         "Fund managers, securities analysts and retail investors who trade LQ45 stocks do not "
         "know whether earnings news is fully priced at the announcement or only gradually "
         "afterwards. They therefore cannot judge whether earnings-surprise information should "
         "enter their post-announcement buy, sell and rebalancing decisions. Regulators (OJK "
         "and the IDX), who must show international investors that the market works, lack "
         "evidence on whether earnings disclosure is efficiently absorbed even in the "
         "market’s leading stocks."],
        ["Research problem",
         "Does the earnings surprise of non-financial LQ45 stocks have a positive effect on "
         "their cumulative abnormal return in the 60 trading days after the announcement "
         "(PEAD)?"],
    ],
    widths=[3.6, 10.4], bold_first_col=True)
T.source("Source: Compiled by the author.")
T.p("Initial evidence that this problem is real comes from prior studies, summarized in "
    "Table 1.2. Drift has been documented in a neighbouring emerging market, including among "
    "its largest firms, and Indonesian studies observe delayed price responses, but none has "
    "tested the earnings-surprise portfolio design within the liquid LQ45 segment.")
T.caption_table("1.2", "Preliminary Evidence on Delayed Price Responses")
T.table(
    ["Study", "Setting", "Preliminary evidence"],
    [
        ["Ho et al. (2024)", "Vietnam, 2009–2020, 20,491 firm-quarters",
         "D5 − D1 spread of 6.5% over 60 days; D1 −2.8%, D5 +3.7%; 5.0% even in the "
         "largest size group"],
        ["Nainggolan (2025)", "IDX, 621 firms, annual earnings releases",
         "Significant abnormal returns in the first week and drift up to 30 days"],
        ["Wijayana & Achjari (2020)", "IDX, IT investment announcements, 2001–2016",
         "Investors need time to recognize value signals; CAR measured from day +2 to +60"],
        ["Ramdhani (2021)", "IDX, 2,531 firm-quarters, 2017–2019",
         "Unexpected cash flows and accruals are related to PEAD; liquidity does not moderate "
         "the relationship"],
        ["Chordia et al. (2009)", "United States, 1972–2005",
         "Drift concentrated in illiquid stocks; transaction costs absorb 70–100% of profit"],
    ],
    widths=[3.6, 4.6, 5.8])
T.source("Source: Compiled from the cited studies.")
T.p("Based on the symptoms and evidence above, the problem of this study is stated as "
    "follows. Although PEAD is documented as a robust anomaly internationally and preliminary "
    "Indonesian evidence suggests delayed price responses to announcements, it is not yet known "
    "whether earnings surprises are followed by a drift among the most liquid, large-"
    "capitalization stocks on the Indonesia Stock Exchange. Without this evidence, investors "
    "cannot determine whether earnings-surprise information remains useful after the "
    "announcement in the segment of the market where it could realistically be traded.")

T.h2("1.3 Research Questions and Research Objectives")
T.p("The problem statement is translated into three research questions and three matching "
    "research objectives, shown in Table 1.3. The first question is the main question of the "
    "study; the second and third describe how the drift occurs.")
T.caption_table("1.3", "Research Questions and Matching Research Objectives")
T.table(
    ["No", "Research question", "Research objective"],
    [
        ["1", "Does earnings surprise have a positive effect on the 60-day post-announcement "
              "cumulative abnormal return, CAR(+2,+61), of non-financial LQ45 stocks?",
         "To test whether earnings surprise has a positive effect on CAR(+2,+61), that is, "
         "whether PEAD exists among non-financial LQ45 stocks."],
        ["2", "Does earnings surprise have a positive effect on the immediate announcement "
              "return, CAR(−1,+1), of non-financial LQ45 stocks?",
         "To test whether the market reacts to the earnings surprise within the three-day "
         "announcement window."],
        ["3", "Does the drift follow the direction of the surprise, with positive drift for "
              "the most positive surprises and negative drift for the most negative "
              "surprises?",
         "To test whether CAR(+2,+61) is positive for the highest-surprise quintile (D5) and "
         "negative for the lowest-surprise quintile (D1)."],
    ],
    widths=[1.0, 6.5, 6.5], align_center_cols=(0,))
T.source("Source: Compiled by the author.")

T.h2("1.4 Scope and Significance of the Study")
T.h3("1.4.1 Scope and Boundaries")
T.p("The study is limited in the following ways, each chosen to keep the test focused and "
    "comparable with Ho et al. (2024):")
T.items([
    "**Object.** Stocks that are LQ45 constituents in the quarter of the earnings "
    "announcement, excluding financial-sector firms, whose accounting and valuation differ "
    "(Cai et al., 2021; Dargenidou et al., 2018).",
    "**Period.** Quarterly earnings announcements for fiscal quarters Q1 2016 to Q4 2025, "
    "about 40 quarters. Fiscal-year 2015 data are used only to compute the first surprises.",
    "**Variables.** Only the direct effect of earnings surprise (X) on PEAD (Y) is tested. "
    "Moderating and control variables used by Ho et al. (2024), such as insider and "
    "institutional ownership, are deliberately excluded and left for future research.",
    "**Measurement.** Earnings surprise uses the seasonal random-walk (time-series) model, "
    "because analyst forecasts are not consistently available for Indonesian stocks over the "
    "whole period. Abnormal returns are measured against the IHSG; LQ45 has too few stocks to "
    "form the size-matched portfolios used by Ho et al. (2024).",
    "**Trading costs.** The study discusses but does not model transaction costs; its "
    "contribution is to show whether drift exists in the low-cost segment of the market.",
], style="alpha")
T.h3("1.4.2 Theoretical Contribution")
T.p("The study adds an out-of-sample test of PEAD to the international literature, which "
    "Hung et al. (2015) describe as still inconclusive. It tests the conjecture in Ho et al. "
    "(2024) that drift may differ in markets with accounting closer to international standards, "
    "and it separates the drift question from the liquidity question raised by Chordia et al. "
    "(2009) by examining only highly liquid stocks. The result will show whether underreaction "
    "and limited-attention explanations of PEAD hold even where information is abundant and "
    "arbitrage is cheap, or whether efficient pricing holds in that segment.")
T.h3("1.4.3 Managerial and Practical Value")
T.items([
    "**Fund managers and securities companies** learn whether an earnings-surprise signal can "
    "improve the timing of post-announcement trades and portfolio rebalancing in LQ45 stocks.",
    "**Retail investors**, the fastest-growing group in the market, learn whether prices of "
    "the most popular stocks keep adjusting after earnings news, which bears on whether buying "
    "or selling immediately after an announcement is reasonable.",
    "**Listed firms** learn how completely the market absorbs their earnings releases, which "
    "is relevant to the timing and clarity of their disclosures.",
    "**The IDX and OJK** obtain evidence on the information efficiency of the market’s most "
    "important segment, which can inform disclosure regulation and investor-education "
    "programs.",
], style="alpha")
T.h3("1.4.4 Societal Impact")
T.p("The study is motivated by the public good as much as by trading. Its societal value "
    "lies in four areas:")
T.items([
    "**Fairness for retail investors.** When prices adjust slowly, investors who trade on "
    "stale prices, typically less-informed individuals, buy too high or sell too low, while "
    "faster and better-informed traders capture the difference. With more than 20 million "
    "investor accounts in Indonesia, about 8.5 million of them holding equities (Irawati, "
    "2025), the "
    "speed at which public information is priced is a question of fairness. Measuring PEAD "
    "shows how large this hidden transfer could be in the most widely held stocks.",
    "**Market credibility and access to capital.** MSCI’s proposed reclassification and the "
    "negative rating outlooks question whether the Indonesian market functions well (MSCI, "
    "2026). A downgrade would push passive foreign funds to sell Indonesian stocks, raising "
    "the cost of capital for listed firms that finance investment and employment. Objective "
    "evidence on information efficiency can help authorities show where the market works and "
    "where reform is still needed.",
    "**Evaluating transparency reforms.** The 2026 reforms (1% shareholder disclosure and a "
    "15% minimum free float) aim to make information more transparent. This study provides "
    "a 2016–2025 baseline of how efficiently earnings news was priced before the reforms, "
    "against which their effect can later be judged.",
    "**Financial literacy.** Clear evidence on whether prices keep moving after earnings news "
    "gives educators and the IDX investor-education programs concrete material on how to read "
    "and use financial reports, instead of relying on rumours or social-media tips.",
], style="alpha")

T.h2("1.5 Research Novelty")
T.p("Table 1.4 positions this study against the closest prior studies. It is the only one "
    "that combines a quarterly earnings-surprise measure, the liquid LQ45 segment, the "
    "standard quintile portfolio design and the most recent period.")
T.caption_table("1.4", "Positioning of This Study against Prior Studies")
T.table(
    ["Study", "Market and sample", "Event and surprise measure", "Liquid segment only",
     "Period"],
    [
        ["Ho et al. (2024)", "Vietnam, all non-financial firms", "Quarterly SUE quintiles",
         "No", "2009–2020"],
        ["Nainggolan (2025)", "IDX, 621 firms", "Annual release; market model, no surprise "
                                                "sorting", "No", "FY 2018"],
        ["Wijayana & Achjari (2020)", "IDX, announcing firms", "IT investment announcements",
         "No", "2001–2016"],
        ["Ramdhani (2021)", "IDX, all listed firms", "Unexpected cash flows and accruals",
         "No", "2017–2019"],
        ["**This study**", "**IDX, non-financial LQ45**", "**Quarterly SUE quintiles**",
         "**Yes**", "**2016–2025**"],
    ],
    widths=[3.2, 3.2, 3.8, 1.8, 2.0], size=9.5, align_center_cols=(3, 4))
T.source("Source: Compiled by the author from the cited studies.")
T.p("Building on this positioning, the study differs from previous research in six "
    "respects:")
T.items([
    "**First earnings-surprise portfolio test of PEAD within the LQ45 index.** Previous "
    "Indonesian studies examine the whole exchange (Nainggolan, 2025; Ramdhani, 2021) or other "
    "types of announcements (Wijayana & Achjari, 2020). To the author’s knowledge, no study has "
    "applied the quintile portfolio design of Bernard and Thomas (1989) and Ho et al. (2024) to "
    "the LQ45 segment alone.",
    "**A tradability lens.** Because LQ45 stocks can be traded in size at low cost, any drift "
    "found here cannot be dismissed as an artifact of illiquidity (Chordia et al., 2009). The "
    "study therefore answers a question that whole-market studies cannot: whether drift exists "
    "where it could actually be exploited.",
    "**An out-of-sample test of Ho et al. (2024).** Ho et al. (2024) explain Vietnamese drift "
    "partly by weak accounting standards and point to Indonesia as a market with standards "
    "closer to IFRS. Testing the same design in Indonesia directly addresses this open "
    "question.",
    "**A recent, longer period.** The 2016–2025 window covers the COVID-19 shock and the "
    "recent surge in retail participation, a period not covered by the studies above.",
    "**Timely evidence for a market under scrutiny.** The study delivers a measure of "
    "information efficiency in Indonesia’s leading stocks at the moment MSCI and the rating "
    "agencies are questioning the market’s transparency (MSCI, 2026). It links an "
    "established finance anomaly to a live policy debate and provides a pre-reform baseline "
    "for evaluating the 2026 transparency reforms.",
    "**Methodological care specific to Indonesia.** Discrete quarterly EPS is derived from "
    "year-to-date interim reports, historical LQ45 constituent lists are used to avoid "
    "survivorship bias, and the quarter, rather than the firm-quarter, is used as the unit of "
    "the portfolio test.",
], style="num")

# ------------------------------------------------------------------------ BAB II
T.page_break()
T.chapter("BAB II", "LITERATURE REVIEW AND THEORETICAL FRAMEWORK")
T.h2("2.1 Literature Synthesis")
T.h3("2.1.1 State of the Art")
T.p("Research on PEAD has moved through three stages. The first stage established that the "
    "drift exists and is robust in the United States (Ball & Brown, 1968; Bernard & Thomas, "
    "1989, 1990). The second stage asked why it exists, testing explanations based on "
    "investor underreaction, limited attention and limits to arbitrage (Chordia et al., 2009; "
    "DellaVigna & Pollet, 2009; Hirshleifer et al., 2009; Mendenhall, 2004). The third and "
    "current stage asks where and when it exists, examining how institutional settings, "
    "ownership structures and accounting standards across countries shape the drift (Cai et "
    "al., 2021; Dargenidou et al., 2018; Ho et al., 2024; Hung et al., 2015). This study "
    "belongs to the third stage. The literature is synthesized below by theme rather than by "
    "author.")
T.h3("2.1.2 Theme 1: Existence and Persistence of PEAD")
T.p("Ball and Brown (1968) showed that the direction of annual earnings changes predicts "
    "returns not only before but also after the announcement. Bernard and Thomas (1989) "
    "formed portfolios on standardized unexpected earnings and found that a strategy long in "
    "the highest decile and short in the lowest earned significant abnormal returns over the "
    "60 trading days after the announcement, and that the drift could not be explained by "
    "risk. Bernard and Thomas (1990) then showed that prices behave as if investors treat "
    "quarterly earnings as a seasonal random walk, ignoring the positive autocorrelation of seasonal "
    "earnings changes; a large part of the drift is realized around the next few earnings "
    "announcements. Fama (1998) concluded that PEAD is one of the few anomalies that survives "
    "reasonable changes in method, and Ball and Brown (2019) found the original pattern "
    "still present fifty years later. Across these studies, the 60-day window after the "
    "announcement, starting on day +2, has become the standard measurement window.")
T.h3("2.1.3 Theme 2: Cross-Country Variation and the Institutional Environment")
T.p("Hung et al. (2015) use the mandatory adoption of IFRS as an information shock and find "
    "that PEAD varies significantly across countries, is insignificant in several of them in "
    "some periods, and declined after countries adopted international accounting standards. "
    "Ho et al. (2024) extend the evidence to Vietnam, a frontier market with weak "
    "accounting standards, concentrated ownership and many individual investors, and document "
    "a 60-day spread of 6.5%. Cai et al. (2021) find that in China institutional ownership is "
    "positively related to PEAD, especially when institutions herd, contrary to the US view "
    "that institutions speed up price discovery. Evidence from Thailand shows that earnings "
    "releases carry mainly firm-specific rather than industry-wide information, while "
    "concentrated ownership and thin analyst coverage shape how that information is processed "
    "(Piyawannasuth, 2025). Together, these studies show that institutional features "
    "matter and that results from one market cannot simply be transferred to another, which "
    "is why Indonesia requires its own test.")
T.h3("2.1.4 Theme 3: Investor Behaviour and Ownership")
T.p("A common interpretation is that investors underreact to earnings news. Hirshleifer et "
    "al. (2009) show that when many firms announce on the same day, the immediate reaction "
    "is weaker and the subsequent drift stronger, consistent with limited attention. "
    "DellaVigna and Pollet (2009) find the same pattern for Friday announcements, when "
    "investors are less attentive. Ownership also matters: Dargenidou et al. (2018) show that "
    "insider trades that confirm the direction of the earnings news allow the drift to "
    "continue, and Ho et al. (2024) find a larger drift in Vietnamese firms with higher insider "
    "ownership but no significant effect of institutional ownership. These findings suggest "
    "that a market with a rapidly growing, retail-dominated investor base, such as Indonesia, "
    "is a plausible setting for underreaction.")
T.h3("2.1.5 Theme 4: Liquidity and Limits to Arbitrage")
T.p("If informed traders could exploit PEAD freely, their trades would remove it. Mendenhall "
    "(2004) shows that the drift is larger for stocks with higher idiosyncratic risk, which "
    "makes arbitrage riskier. Chordia et al. (2009) find that PEAD is concentrated in illiquid "
    "stocks and that trading costs absorb most of the profit, so the anomaly persists because "
    "it is costly to trade. In Indonesia, Ramdhani (2021) finds that liquidity, measured by "
    "the Amivest ratio, does not moderate the relationship between the components of earnings "
    "surprise and PEAD. The theme implies a clear test: in highly liquid stocks, limits to "
    "arbitrage are weakest, so finding drift there would be strong evidence of "
    "underreaction, while finding none would be consistent with efficient pricing where "
    "arbitrage is cheap.")
T.h3("2.1.6 Theme 5: Measuring Earnings Surprise and Abnormal Returns")
T.p("Earnings surprise can be measured against analyst forecasts or against a time-series "
    "expectation. Livnat and Mendenhall (2006) show that both measures produce a significant "
    "drift and that the time-series measure, defined as the seasonal change in EPS scaled by "
    "price, is a valid alternative where forecasts are unavailable. Ho et al. (2024) adopt "
    "this measure for Vietnam for that reason. Abnormal returns are commonly computed as the "
    "difference between a stock’s return and a benchmark return, and simple market-adjusted "
    "returns perform well in daily event studies (Brown & Warner, 1985). Ho et al. (2024) use "
    "size-adjusted returns as the main measure and market-index-adjusted returns as a "
    "robustness check, obtaining similar results. In thinly traded markets, beta estimates "
    "can be biased, which led Wijayana and Achjari (2020) to use adjusted betas; this concern "
    "is minor for LQ45 stocks, which trade every day.")
T.h3("2.1.7 Theme 6: Indonesian Evidence")
T.p("Nainggolan (2025) applies the market model to annual earnings releases of 621 IDX firms "
    "and finds no information leakage before the release but significant abnormal returns in "
    "the first week after it and drift up to 30 days, concluding that the IDX is not "
    "efficient in the semi-strong form. Wijayana and Achjari (2020) study information "
    "technology investment announcements from 2001 to 2016 and find that the market reaction "
    "occurs after, not around, the announcement. Ramdhani (2021) decomposes earnings surprise "
    "into unexpected cash flows and unexpected accruals and finds that both are related to "
    "PEAD for IDX firms over 2017–2019. All three studies point to delayed price responses in "
    "Indonesia, but each uses a broad sample, a different event type or a different surprise "
    "measure.")
T.h3("2.1.8 Summary of Prior Studies")
T.p("Table 2.1 summarizes the main empirical studies reviewed, their methods and findings, and "
    "how this study relates to each.")
T.caption_table("2.1", "Summary of Prior Empirical Studies")
T.table(
    ["No", "Author (year)", "Context and method", "Key finding", "Relation to this study"],
    [
        ["1", "Ball & Brown (1968)", "US; annual earnings; abnormal-return index",
         "Prices drift in the direction of the earnings change", "Origin of the anomaly"],
        ["2", "Bernard & Thomas (1989)", "US; SUE decile portfolios; 60-day CAR",
         "Significant long-short drift not explained by risk", "Portfolio design and window"],
        ["3", "Bernard & Thomas (1990)", "US; seasonal earnings autocorrelation",
         "Investors ignore implications of current for future earnings",
         "Theoretical mechanism for H1"],
        ["4", "Livnat & Mendenhall (2006)", "US; analyst vs time-series surprise",
         "Time-series SUE yields significant drift", "Justifies SUE measure"],
        ["5", "Chordia et al. (2009)", "US; liquidity-sorted portfolios",
         "Drift concentrated in illiquid stocks; costs absorb 70–100% of profit",
         "Motivates LQ45 focus"],
        ["6", "Hirshleifer et al. (2009)", "US; same-day announcements",
         "More competing news → weaker reaction, stronger drift", "Limited attention (H1, H2)"],
        ["7", "Hung et al. (2015)", "International; IFRS adoption as shock",
         "PEAD varies across countries; insignificant in some", "Need for country test"],
        ["8", "Dargenidou et al. (2018)", "UK; insider trading and PEAD",
         "Confirmatory insider trades prolong drift", "Ownership channel (excluded)"],
        ["9", "Cai et al. (2021)", "China; institutional ownership and herding",
         "Institutions herding increases PEAD", "Emerging-market setting"],
        ["10", "Ho et al. (2024)", "Vietnam 2009–2020; SUE quintiles; t-test and regression",
         "6.5% 60-day spread; larger with insider ownership", "Main reference design"],
        ["11", "Wijayana & Achjari (2020)", "Indonesia 2001–2016; IT investment events",
         "Reaction occurs after the announcement", "Delayed response in Indonesia"],
        ["12", "Nainggolan (2025)", "Indonesia; 621 firms; market model",
         "Abnormal returns in first week; drift up to 30 days", "Whole-market IDX evidence"],
        ["13", "Ramdhani (2021)", "Indonesia 2017–2019; 2,531 firm-quarters",
         "Surprise components related to PEAD; liquidity not a moderator",
         "Local thesis evidence"],
    ],
    widths=[0.8, 2.8, 3.6, 3.8, 3.0], size=9.5, align_center_cols=(0,))
T.source("Source: Compiled by the author from the cited studies.")
T.h3("2.1.9 Research Gap")
T.p("The synthesis reveals five gaps. First, an **empirical gap**: Ho et al. (2024) leave "
    "open whether PEAD exists in a market with accounting closer to international standards, "
    "and no study has tested the earnings-surprise portfolio design among LQ45 stocks. "
    "Second, a **tradability gap**: Indonesian evidence comes from whole-market samples in "
    "which drift may be confined to illiquid stocks (Chordia et al., 2009), so it is unknown "
    "whether drift exists where it could be traded. Third, a **methodological gap**: existing "
    "Indonesian studies use annual releases (Nainggolan, 2025), non-earnings events (Wijayana "
    "& Achjari, 2020) or decomposed surprise measures (Ramdhani, 2021), rather than the "
    "standard quarterly SUE quintile spread that allows comparison with international results. "
    "Fourth, a **temporal gap**: no study covers the 2016–2025 period, which includes the "
    "COVID-19 shock and the recent surge in retail participation. Fifth, a **policy-relevance "
    "gap**: although MSCI and the rating agencies question the transparency of the Indonesian "
    "market (MSCI, 2026), no study provides recent evidence on how efficiently earnings news "
    "is priced in its leading stocks. This study addresses all five gaps by applying the design of Ho et al. (2024) to non-financial LQ45 stocks over "
    "2016–2025.")

T.h2("2.2 Theoretical Framework and Conceptual Model")
T.h3("2.2.1 Efficient Market Hypothesis")
T.p("Fama (1970) defines a market as efficient in the semi-strong form when prices fully "
    "reflect all publicly available information. Under this view, an earnings surprise is "
    "absorbed into the price at the announcement, so the surprise should predict the "
    "announcement return but not returns in the following weeks. The efficient market "
    "hypothesis therefore provides the null expectation of this study: no relation between "
    "earnings surprise and CAR(+2,+61).")
T.h3("2.2.2 Investor Underreaction")
T.p("Bernard and Thomas (1990) argue that investors fail to recognize that seasonal earnings "
    "changes are positively autocorrelated, so they underestimate how much a surprise this "
    "quarter predicts surprises in later quarters. Prices therefore adjust only partly at the "
    "announcement and continue adjusting as later information arrives. Barberis et al. "
    "(1998) give a behavioural basis for this pattern: investors exhibit conservatism, "
    "updating their beliefs too slowly in response to new public information such as "
    "earnings. Underreaction predicts that the surprise is positively related to both the "
    "announcement return and the subsequent drift.")
T.h3("2.2.3 Limited Attention")
T.p("Investors have limited time and cognitive capacity. When attention is scarce, because "
    "many firms announce at once or because announcements fall on low-attention days, the "
    "initial reaction is weaker and more of the adjustment happens later (DellaVigna & Pollet, "
    "2009; Hirshleifer et al., 2009). In Indonesia, the clustering of releases near reporting "
    "deadlines and the large number of new retail investors make limited attention a plausible "
    "source of delayed adjustment.")
T.h3("2.2.4 Limits to Arbitrage")
T.p("Shleifer and Vishny (1997) explain that mispricing can persist because arbitrage is "
    "costly and risky. For PEAD, the drift persists where trading costs and idiosyncratic risk "
    "are high (Chordia et al., 2009; Mendenhall, 2004). This theory explains why the choice of "
    "LQ45 matters: liquidity lowers the cost of arbitrage, so the theory predicts a smaller "
    "drift among these stocks. The hypotheses below follow the underreaction and "
    "limited-attention predictions; limits-to-arbitrage theory explains why the test in LQ45 "
    "is demanding and why a non-significant result would still be informative.")
T.h3("2.2.5 Information Environment and Transparency")
T.p("The speed at which prices absorb earnings news also depends on the quality of the "
    "information environment. Drift declined after countries adopted international "
    "accounting standards (Hung et al., 2015) and is stronger where ownership is "
    "concentrated among insiders, because concentrated ownership makes earnings more opaque "
    "and harder for outside investors to interpret (Ho et al., 2024). Insiders whose trades "
    "confirm the earnings news can also prolong the drift (Dargenidou et al., 2018). This "
    "perspective connects PEAD to the 2026 debate on Indonesian market transparency: concerns "
    "about opaque shareholding structures (MSCI, 2026) are, in theoretical terms, concerns "
    "about an information environment that slows the incorporation of public news into "
    "prices. A drift among LQ45 stocks would indicate that such frictions affect even the "
    "market’s most visible firms.")
T.h3("2.2.6 Integrating the Theories")
T.p("The theories make competing predictions, which is what makes the test informative. "
    "Table 2.2 summarizes what each theory predicts for the relation between earnings surprise "
    "and post-announcement returns among LQ45 stocks.")
T.caption_table("2.2", "Theoretical Predictions for Earnings Surprise and PEAD")
T.table(
    ["Theory", "Core argument", "Prediction for LQ45 stocks", "Hypotheses"],
    [
        ["Efficient market hypothesis (Fama, 1970)",
         "Public information is priced fully and immediately",
         "Reaction at announcement only; no drift", "Null of H1; supports H2"],
        ["Underreaction (Barberis et al., 1998; Bernard & Thomas, 1990)",
         "Investors update too slowly to earnings news",
         "Reaction at announcement plus drift in the same direction", "H1, H2, H3"],
        ["Limited attention (Hirshleifer et al., 2009)",
         "Scarce attention delays processing of news",
         "Weaker initial reaction, stronger drift", "H1, H3"],
        ["Limits to arbitrage (Chordia et al., 2009; Shleifer & Vishny, 1997)",
         "Costly arbitrage lets mispricing persist",
         "Smaller drift in liquid stocks than in illiquid stocks", "Explains a null H1"],
        ["Information environment (Ho et al., 2024; Hung et al., 2015)",
         "Opacity slows interpretation of earnings",
         "Drift possible even in large stocks if opacity is high", "H1, H3"],
    ],
    widths=[3.6, 3.6, 4.0, 2.8], size=9.5)
T.source("Source: Compiled by the author.")
T.h3("2.2.7 Conceptual Model")
T.p("Figure 2.1 shows the conceptual model. The independent variable is the earnings surprise "
    "(X), with two dimensions: its direction (positive or negative) and its magnitude. The "
    "dependent variable is post-earnings announcement drift (Y), measured mainly by "
    "CAR(+2,+61). The model also tests the announcement reaction, CAR(−1,+1), and the "
    "direction of the drift in the extreme quintiles.")
T.figure(figdir / "fig_conceptual_model.png", "2.1", "Conceptual Model of the Study",
         width_cm=14.5, source="Source: Developed by the author based on Bernard and Thomas "
         "(1989) and Ho et al. (2024).")
T.p("Table 2.3 identifies the role of each variable. The model has no mediating or "
    "moderating variable by design: the purpose of the study is to establish whether PEAD "
    "exists in LQ45 before examining what drives it, which follows the first step of Ho et "
    "al. (2024). Quarter fixed effects in the regression test control for shocks that affect "
    "all stocks in the same quarter, such as the COVID-19 crash in 2020.")
T.caption_table("2.3", "Identification of Research Variables")
T.table(
    ["Role", "Variable", "Measure", "Basis"],
    [
        ["Independent (IV)", "Earnings surprise (X)", "SUE; quintile rank SUER",
         "Bernard & Thomas (1989); Livnat & Mendenhall (2006)"],
        ["Dependent (DV)", "Post-earnings announcement drift (Y)",
         "CAR(+2,+61); supplementary CAR(+2,+31)", "Bernard & Thomas (1989); Ho et al. (2024)"],
        ["Dependent (H2)", "Announcement reaction", "CAR(−1,+1)", "Ho et al. (2024)"],
        ["Mediating", "None (by design)", "–", "Existence test precedes mechanism tests"],
        ["Moderating", "None (by design)", "–",
         "Ownership and liquidity left for future research"],
        ["Control", "Quarter fixed effects", "Quarter dummies", "Ho et al. (2024)"],
    ],
    widths=[2.9, 3.6, 3.6, 3.9], bold_first_col=True)
T.source("Source: Compiled by the author.")

T.h2("2.3 Hypotheses Development")
T.h3("2.3.1 Earnings Surprise and Post-Earnings Announcement Drift (H1)")
T.p("Underreaction theory predicts that prices adjust only partly to an earnings surprise at "
    "the announcement (Barberis et al., 1998; Bernard & Thomas, 1990), and limited attention "
    "slows the adjustment further (Hirshleifer et al., 2009). Stocks with the most positive "
    "surprises should therefore continue to rise and stocks with the most negative surprises "
    "should continue to fall. Empirically, this pattern holds in the United States (Bernard & "
    "Thomas, 1989), across many countries (Hung et al., 2015), in Vietnam including its largest "
    "firms (Ho et al., 2024), and, according to preliminary studies, on the IDX (Nainggolan, "
    "2025; Ramdhani, 2021). Although LQ45 stocks are liquid, the retail-dominated investor base "
    "and clustered reporting dates suggest that underreaction may persist. The main hypothesis "
    "is therefore:")
T.p("**H1:** Earnings surprise has a positive effect on post-earnings announcement drift; the "
    "CAR(+2,+61) of the highest-surprise quintile (D5) is greater than that of the "
    "lowest-surprise quintile (D1).", indent=0)
T.h3("2.3.2 Earnings Surprise and the Announcement Reaction (H2)")
T.p("Even when the adjustment is incomplete, a part of the earnings news should be priced at "
    "the announcement. Under both efficient-market and underreaction views, a positive surprise "
    "raises the stock price and a negative surprise lowers it in the days around the "
    "announcement; the two views differ only in whether the reaction is complete. Ho et al. "
    "(2024) find a 3-day spread of 3.0% in Vietnam, and Nainggolan (2025) finds significant "
    "abnormal returns in the first days after IDX earnings releases. A significant announcement "
    "reaction also confirms that SUE captures news that investors regard as value-relevant. "
    "Hence:")
T.p("**H2:** Earnings surprise has a positive effect on the announcement reaction; the "
    "CAR(−1,+1) of D5 is greater than that of D1.", indent=0)
T.h3("2.3.3 Direction of the Drift (H3)")
T.p("PEAD is defined as a drift in the direction of the surprise. A positive spread between D5 "
    "and D1 could, in principle, arise entirely from one side, for example if only bad news "
    "were priced slowly. Underreaction implies that both sides drift: good news continues to be "
    "incorporated upwards and bad news downwards. Ho et al. (2024) report a 60-day CAR of "
    "+3.7% for D5 and −2.8% for D1. Testing each side separately matches the two dimensions of "
    "the earnings surprise, direction and magnitude, and shows whether the drift is symmetric. "
    "Hence:")
T.p("**H3a:** The CAR(+2,+61) of the highest-surprise quintile (D5) is positive.", indent=0)
T.p("**H3b:** The CAR(+2,+61) of the lowest-surprise quintile (D1) is negative.", indent=0)
T.p("All hypotheses are directional and are therefore tested with one-tailed tests. Table 2.4 "
    "summarizes them.")
T.caption_table("2.4", "Summary of Hypotheses")
T.table(
    ["Hypothesis", "Relationship", "Test statistic", "Expected sign"],
    [
        ["H1 (main)", "SUE → CAR(+2,+61)",
         "Mean quarterly D5 − D1 spread; α_{1} in regression", "Positive"],
        ["H2", "SUE → CAR(−1,+1)", "Mean quarterly D5 − D1 spread; α_{1} in regression",
         "Positive"],
        ["H3a", "Positive surprise → CAR(+2,+61)", "Mean CAR(+2,+61) of D5", "Positive"],
        ["H3b", "Negative surprise → CAR(+2,+61)", "Mean CAR(+2,+61) of D1", "Negative"],
    ],
    widths=[2.4, 4.0, 5.0, 2.6], align_center_cols=(3,))
T.source("Source: Compiled by the author.")

# ------------------------------------------------------------------------ BAB III
T.page_break()
T.chapter("BAB III", "RESEARCH METHODOLOGY")
T.h2("3.1 Research Design")
T.p("This study is a quantitative, hypothesis-testing study that uses an archival event-study "
    "design (Brown & Warner, 1985; Fama et al., 1969). Its purpose is explanatory: it tests "
    "whether earnings surprise predicts subsequent abnormal returns, as stated in H1 to H3. "
    "The researcher does not manipulate any variable; the data are generated by actual trading "
    "on the IDX. The design follows Ho et al. (2024) so that the results can be compared "
    "directly with evidence from Vietnam. Table 3.1 summarizes the design elements in the "
    "terms used by Sekaran and Bougie (2016).")
T.caption_table("3.1", "Research Design Elements")
T.table(
    ["Element", "Choice in this study", "Reason"],
    [
        ["Purpose", "Hypothesis testing (explanatory, causal-predictive)",
         "Tests directional hypotheses derived from theory"],
        ["Researcher interference", "Minimal; no manipulation",
         "Returns and earnings are observed, not induced"],
        ["Study setting", "Non-contrived field setting (the actual IDX market)",
         "Real prices reflect real investor decisions"],
        ["Strategy", "Archival event study using secondary data",
         "Standard method for price reactions to disclosures"],
        ["Time horizon", "Longitudinal: 40 quarters, Q1 2016 – Q4 2025",
         "The drift test requires repeated quarterly observations"],
        ["Unit of analysis", "Firm-quarter (one earnings announcement); the quarter for the "
                             "portfolio t-test", "Each announcement is one event"],
        ["Data", "Quantitative secondary data from the IDX", "Public, objective and replicable"],
    ],
    widths=[3.2, 5.6, 5.2], bold_first_col=True)
T.source("Source: Compiled by the author based on Sekaran and Bougie (2016).")
T.p("Figure 3.1 shows the research procedure, from defining the universe to the robustness "
    "checks.")
T.figure(figdir / "fig_research_procedure.png", "3.1", "Research Procedure", width_cm=12.5,
         source="Source: Developed by the author based on Ho et al. (2024).")

T.h2("3.2 Operational Definition and Measurement")
T.h3("3.2.1 Operational Definitions")
T.p("Each variable is defined by combining two sources and then stated in measurable terms, "
    "as shown in Table 3.2.")
T.caption_table("3.2", "Operational Definitions of Variables")
T.table(
    ["Variable", "Conceptual definition", "Operational definition", "Reference"],
    [
        ["Earnings surprise (X)",
         "The part of reported earnings that was not expected by the market.",
         "EPS of quarter t minus EPS of the same quarter one year earlier (t−4), divided by "
         "the share price at the end of quarter t; ranked into quintiles each quarter (SUER).",
         "Ho et al. (2024); Livnat & Mendenhall (2006)"],
        ["Post-earnings announcement drift (Y)",
         "The tendency of prices to keep moving in the direction of the earnings surprise "
         "after the announcement, because investors do not fully understand the implications "
         "of current earnings.",
         "Cumulative abnormal return from trading day +2 to +61 after the announcement; "
         "daily abnormal return = stock return − IHSG return.",
         "Bernard & Thomas (1989); Ho et al. (2024)"],
        ["Announcement reaction",
         "The immediate price response to the earnings release.",
         "Cumulative abnormal return from trading day −1 to +1 around the announcement.",
         "Ho et al. (2024)"],
    ],
    widths=[2.8, 3.7, 4.7, 2.8], size=9.5, bold_first_col=True)
T.source("Source: Compiled by the author.")
T.h3("3.2.2 Operationalization Matrix")
T.p("Because the study uses secondary data, each concept is broken down into dimensions and "
    "indicators that are calculated from financial statements and prices rather than "
    "questionnaire items. Table 3.3 presents the operationalization matrix and the measurement "
    "scale of each indicator.")
T.caption_table("3.3", "Operationalization Matrix")
T.table(
    ["Variable", "Dimension", "Indicator", "Scale", "Data source"],
    [
        ["Earnings surprise (X)", "D1: Direction of surprise",
         "Sign of EPS_{t} − EPS_{t−4} (positive / negative)", "Nominal",
         "Quarterly financial statements"],
        ["", "D2: Magnitude of surprise",
         "SUE = (EPS_{t} − EPS_{t−4}) / P_{t}", "Ratio",
         "Quarterly financial statements; IDX prices"],
        ["", "D2: Relative magnitude",
         "SUER: quintile rank of SUE within the quarter (1 = lowest, 5 = highest)",
         "Ordinal", "Calculated"],
        ["Post-earnings announcement drift (Y)", "D1: Immediate reaction",
         "CAR(−1,+1)", "Ratio", "IDX daily closing prices; IHSG"],
        ["", "D2: 30-day drift", "CAR(+2,+31)", "Ratio", "IDX daily closing prices; IHSG"],
        ["", "D3: 60-day drift (main)", "CAR(+2,+61)", "Ratio",
         "IDX daily closing prices; IHSG"],
        ["", "Portfolio drift", "Quarterly spread: mean CAR of D5 − mean CAR of D1", "Ratio",
         "Calculated"],
    ],
    widths=[2.8, 2.8, 4.2, 1.6, 2.6], size=9.5)
T.source("Source: Compiled by the author based on Ho et al. (2024).")
T.h3("3.2.3 Measurement Formulas")
T.p("Indonesian interim financial statements report year-to-date figures. Discrete quarterly "
    "EPS is therefore derived by subtracting the previous cumulative figure:")
T.equation("EPS_{q} = EPS^{YTD}_{q} − EPS^{YTD}_{q−1}   (EPS_{Q1} = EPS^{YTD}_{Q1})", 1)
T.p("where EPS^{YTD}_{q} is cumulative EPS up to quarter q of the same fiscal year; for "
    "example, Q4 EPS equals full-year EPS minus nine-month EPS. The earnings surprise follows "
    "Livnat and Mendenhall (2006) and Ho et al. (2024):", indent=0)
T.equation("SUE_{i,t} = (EPS_{i,t} − EPS_{i,t−4}) / P_{i,t}", 2)
T.p("where EPS_{i,t} is the EPS of firm i in quarter t, EPS_{i,t−4} is the EPS of the same "
    "quarter one year earlier, adjusted for stock splits, bonus shares and rights issues so "
    "that it is on the same per-share basis, and P_{i,t} is the closing price on the last "
    "trading day of quarter t. Each quarter, firms are sorted into five quintiles by SUE, and "
    "SUER takes values 1 (D1, most negative) to 5 (D5, most positive). Daily returns and "
    "abnormal returns are:", indent=0)
T.equation("R_{i,d} = (P_{i,d} − P_{i,d−1}) / P_{i,d−1};    AR_{i,d} = R_{i,d} − R_{IHSG,d}", 3)
T.p("where P_{i,d} is the closing price adjusted for corporate actions and dividends where "
    "data allow, and R_{IHSG,d} is the daily return of the IHSG. The cumulative abnormal "
    "return over the window from day a to day b relative to the announcement date (day 0) "
    "is:", indent=0)
T.equation("CAR_{i}(a,b) = Σ_{d=a}^{b} AR_{i,d}", 4)
T.p("The study uses three windows: (−1,+1), (+2,+31) and (+2,+61). Days 0 and +1 are "
    "excluded from the drift windows so that the drift is measured after the announcement "
    "reaction is complete (Bernard & Thomas, 1989). The quarterly PEAD spread is:", indent=0)
T.equation("Spread_{q} = mean CAR(+2,+61)_{D5,q} − mean CAR(+2,+61)_{D1,q}", 5)
T.p("All continuous variables are winsorized at the 1st and 99th percentiles to limit the "
    "influence of outliers (Ho et al., 2024).", indent=0)
T.h3("3.2.4 Measurement Instruments, Validity and Reliability")
T.p("No questionnaire or rating scale is used, so Likert scales and Cronbach’s alpha do not "
    "apply. Instead, the quality of measurement is ensured in the following ways.")
T.items([
    "**Construct (content) validity.** SUE and CAR are the established measures of earnings "
    "surprise and drift in the PEAD literature (Bernard & Thomas, 1989; Ho et al., 2024; "
    "Livnat & Mendenhall, 2006), so they measure the intended constructs.",
    "**Criterion validity.** If SUE captures value-relevant news, it must be positively related "
    "to the announcement return. The H2 test therefore also serves as a validity check of the "
    "surprise measure.",
    "**Robustness of the abnormal-return measure.** Results will be re-estimated with the LQ45 "
    "index as the benchmark, because a measurement choice should not drive the conclusion "
    "(Ho et al., 2024).",
    "**Reliability of the data.** A random 10% of firm-quarters will be checked by hand: EPS "
    "against the original financial statements on the IDX website and prices against a second "
    "data provider. Any discrepancy is resolved using the original IDX document, and the "
    "error rate is reported.",
    "**Reliability of the computation.** All calculations are scripted (Python or Stata), so "
    "rerunning the code on the same data reproduces every number. The code and data log are "
    "archived as thesis appendices.",
], style="alpha")

T.h2("3.3 Population and Sampling Design")
T.p("The study uses non-probability judgment (purposive) sampling, and the unit of observation "
    "is the firm-quarter. The sample is determined in five steps.")
T.p("**Step 1: Define the population.** The population is all stocks that were constituents "
    "of the LQ45 index on the IDX at any time during 2016–2025. LQ45 is chosen because it "
    "contains highly liquid stocks, so daily returns for calculating CAR are available on "
    "every trading day.")
T.p("**Step 2: Determine the sampling frame.** The sampling frame is the list of LQ45 "
    "constituents announced by the IDX at each review (February and August). Historical lists "
    "are used rather than the latest list, so firms that left the index are still covered for "
    "the periods in which they were members.")
T.p("**Step 3: Determine the sampling design.** Judgment sampling is applied with criteria "
    "that follow Ho et al. (2024):")
T.items([
    "the stock is an LQ45 constituent in the quarter of the earnings announcement;",
    "the firm is not in the financial sector, because its accounting (for example, "
    "mark-to-market accounting) and valuation differ (Cai et al., 2021; Dargenidou et al., "
    "2018; Ho et al., 2024);",
    "the firm has financial statements for eight consecutive quarters, so that EPS_{t−4} and "
    "the derived quarterly EPS are available; and",
    "trading data are available to compute CAR from day −1 to day +61.",
], style="alpha")
T.p("**Step 4: Determine the sample size.** The sample size is not set by a formula; all "
    "firm-quarters that meet the criteria are used. With about 37–40 non-financial "
    "constituents per period over 40 quarters, the maximum is about 1,500–1,600 firm-quarter "
    "observations before filtering, and each quintile holds about seven to eight stocks per "
    "quarter. This size is adequate for two reasons. First, the portfolio t-test uses about "
    "40 quarterly spreads, above the minimum of 30 that Sekaran and Bougie (2016) recommend "
    "for most research; with 40 observations, a one-tailed test at α = 0.05 has 80% power to "
    "detect a mean spread of about 0.39 standard deviations of the quarterly spread. Second, "
    "the regression test uses all firm-quarters, which gives ample degrees of freedom for a "
    "model with one variable of interest and quarter fixed effects.")
T.p("**Step 5: Sampling process.** Price data and financial statements are collected for each "
    "stock and quarter. The number of observations excluded at each criterion is recorded in a "
    "sample selection table, following Table 1 of Ho et al. (2024). Table 3.4 shows the format "
    "that will be completed in Chapter IV.")
T.caption_table("3.4", "Sample Selection Procedure")
T.table(
    ["Step", "Criterion", "Firm-quarters"],
    [
        ["1", "LQ45 constituent firm-quarters, Q1 2016 – Q4 2025 (45 × 40)", "1,800"],
        ["2", "Less: financial-sector firm-quarters", "(…)"],
        ["3", "Less: firm-quarters without eight consecutive quarters of statements", "(…)"],
        ["4", "Less: firm-quarters without trading data for CAR(−1,+61)", "(…)"],
        ["", "**Final sample**", "**…**"],
    ],
    widths=[1.2, 9.8, 3.0], align_center_cols=(0, 2))
T.source("Source: Format adapted from Ho et al. (2024); figures to be reported in Chapter IV.")

T.h2("3.4 Data Collection Methods and Bias Reduction")
T.p("The study uses documentation of secondary data only; no questionnaire, interview or "
    "observation of people is involved. Table 3.5 lists the data, their sources and the "
    "collection protocol.")
T.caption_table("3.5", "Data Sources and Collection Protocol")
T.table(
    ["Data", "Source", "Collection protocol"],
    [
        ["LQ45 constituent lists, 2016–2025", "IDX announcements (www.idx.co.id)",
         "Download each February and August review; record effective dates"],
        ["Quarterly financial statements (EPS, shares outstanding)",
         "IDX listed-company disclosure pages", "Record cumulative EPS per report; derive "
                                                "quarterly EPS with Equation (1)"],
        ["Earnings announcement date", "IDX disclosure system (keterbukaan informasi)",
         "Date the statement is first published; if published after market close, day 0 is "
         "the next trading day"],
        ["Daily closing prices and IHSG", "IDX trading data; cross-checked with a second "
                                          "provider", "Adjust for splits, bonus and rights "
                                                      "issues"],
        ["Corporate actions", "IDX announcements", "Build adjustment factors for EPS and "
                                                   "prices"],
    ],
    widths=[3.9, 4.3, 5.8], size=9.5)
T.source("Source: Compiled by the author.")
T.p("Biases in archival studies differ from survey biases but are just as serious. Table 3.6 "
    "lists the main risks and the steps taken to minimize each.")
T.caption_table("3.6", "Potential Biases and Mitigation Steps")
T.table(
    ["Bias", "Risk", "Mitigation"],
    [
        ["Survivorship bias", "Using today’s LQ45 list keeps only successful firms",
         "Use historical constituent lists for every review period"],
        ["Look-ahead bias", "Using information not yet public when the portfolio is formed",
         "Rank SUE only within the announcement quarter; use actual publication dates"],
        ["Event-date error", "Wrong day 0 shifts returns between windows",
         "Use first IDX publication time; after-hours releases move to the next day"],
        ["Measurement error in EPS", "Year-to-date reports and corporate actions distort EPS",
         "Derive quarterly EPS (Equation 1); adjust EPS_{t−4} for share changes"],
        ["Outliers", "A few extreme values drive the mean",
         "Winsorize at 1% and 99%; report medians; Wilcoxon test"],
        ["Cross-sectional dependence", "Announcements cluster in the same weeks",
         "Use the quarter as the unit of the t-test; quarter fixed effects; clustered errors"],
        ["Data-snooping bias", "Choosing windows after seeing results",
         "Windows and hypotheses fixed in this proposal, following prior studies"],
        ["Selection bias", "Excluding financial firms limits generality",
         "State the boundary explicitly; conclusions apply to non-financial LQ45 stocks"],
    ],
    widths=[3.1, 4.9, 6.0], size=9.5, bold_first_col=True)
T.source("Source: Compiled by the author.")

T.h2("3.5 Data Analysis Plan")
T.h3("3.5.1 Data Screening and Cleaning")
T.p("Before testing, the data are screened by (a) checking for missing EPS, prices and "
    "announcement dates and excluding firm-quarters that fail the sample criteria; (b) "
    "verifying that derived quarterly EPS sums to annual EPS for each fiscal year; (c) "
    "checking for duplicate announcements and for implausible daily returns caused by "
    "unadjusted corporate actions; and (d) winsorizing SUE and CAR at the 1st and 99th "
    "percentiles.")
T.h3("3.5.2 Descriptive Statistics")
T.p("Descriptive statistics (number of observations, mean, median, standard deviation, "
    "minimum and maximum) are reported for SUE and for CAR in each window, overall and by "
    "quintile. The average cumulative abnormal return path of each quintile from day −1 to "
    "day +61 is plotted, which shows visually whether the extreme quintiles diverge after the "
    "announcement. The number of stocks per quintile per quarter is also reported.")
T.h3("3.5.3 Inferential Tests")
T.p("**Test A, portfolio test (H1, H2, H3).** For each quarter, the mean CAR of every quintile "
    "and the D5 − D1 spread are computed (Equation 5). The time series of about 40 quarterly "
    "spreads is tested against zero with a one-sample t-test, following the portfolio test of "
    "Ho et al. (2024); using the quarter as the unit of observation accounts for the "
    "correlation among stocks that announce in the same period. For H1 "
    "the test uses CAR(+2,+61), for H2 CAR(−1,+1), and for H3a and H3b the quarterly mean "
    "CAR(+2,+61) of D5 and D1, respectively. Because quarterly spreads may be autocorrelated, "
    "Newey-West standard errors are also reported, and the Wilcoxon signed-rank test is used "
    "as a non-parametric check given the moderate number of quarters.")
T.p("**Test B, regression test (H1, H2).** All firm-quarters are pooled in the following "
    "regression:", indent=0)
T.equation("CAR_{i,q} = α_{0} + α_{1} SUER′_{i,q} + Σ γ_{q} Quarter_{q} + ε_{i,q}", 6)
T.p("where CAR_{i,q} is CAR(+2,+61) for H1 or CAR(−1,+1) for H2, SUER′ = (SUER − 1)/4 rescales "
    "the quintile rank to the 0–1 range so that α_{1} equals the return difference between "
    "D5 and D1, the rank-scaling convention used in this literature (Cai et al., 2021; Hung "
    "et al., 2015), and Quarter_{q} are quarter fixed "
    "effects. Standard errors are clustered by firm because the same firm appears in many "
    "quarters (Petersen, 2009). No control variables are added, consistent with the scope of "
    "the study.", indent=0)
T.p("Significance is evaluated at α = 0.05 with one-tailed tests, because all hypotheses "
    "are directional; results at the 1% and 10% levels are also reported. Table 3.7 summarizes "
    "the techniques and decision rules.")
T.caption_table("3.7", "Hypothesis Testing Techniques and Decision Rules")
T.table(
    ["Hypothesis", "Technique", "Decision rule (α = 0.05, one-tailed)"],
    [
        ["H1", "t-test on quarterly CAR(+2,+61) spread; regression (6)",
         "Supported if mean spread > 0 and α_{1} > 0 with p < 0.05"],
        ["H2", "t-test on quarterly CAR(−1,+1) spread; regression (6)",
         "Supported if mean spread > 0 and α_{1} > 0 with p < 0.05"],
        ["H3a", "t-test on quarterly mean CAR(+2,+61) of D5",
         "Supported if mean > 0 with p < 0.05"],
        ["H3b", "t-test on quarterly mean CAR(+2,+61) of D1",
         "Supported if mean < 0 with p < 0.05"],
    ],
    widths=[2.2, 6.0, 5.8], align_center_cols=(0,))
T.source("Source: Compiled by the author.")
T.p("Because each quintile holds only about seven to eight stocks per quarter, a weak or "
    "non-significant result is possible. Such a result is still informative: it would indicate "
    "that earnings news is priced quickly in the most liquid segment of the IDX, consistent "
    "with limits-to-arbitrage theory (Chordia et al., 2009).")
T.h3("3.5.4 Robustness Checks")
T.p("Table 3.8 lists the robustness checks planned to show that the conclusions do not depend "
    "on particular measurement choices.")
T.caption_table("3.8", "Planned Robustness Checks")
T.table(
    ["Check", "Purpose"],
    [
        ["LQ45 index as the benchmark instead of the IHSG",
         "Benchmark that matches the large-cap sample"],
        ["Terciles instead of quintiles", "Larger portfolios (about 12–13 stocks) reduce noise"],
        ["30-day window, CAR(+2,+31)", "Shows how much of the drift occurs in the first month"],
        ["Excluding 2020 quarters", "Removes the COVID-19 crash and recovery"],
        ["Median spread and Wilcoxon signed-rank test", "Limits the influence of outliers"],
        ["Supplementary: Q1–Q3 2026 announcements (descriptive)",
         "Shows whether the drift pattern changed during the 2026 MSCI and rating-agency "
         "scrutiny; too few quarters for formal tests"],
    ],
    widths=[6.5, 7.5])
T.source("Source: Compiled by the author.")
T.h3("3.5.5 Software and Qualitative Analysis")
T.p("Data are processed in Microsoft Excel and analysed in Python (pandas and statsmodels) or "
    "Stata. Qualitative analysis is not applicable, because the study uses only quantitative "
    "secondary data.")

T.h2("3.6 Research Ethics and Thesis Timeline")
T.h3("3.6.1 Research Ethics")
T.p("The study does not involve human participants, so informed consent and participant "
    "confidentiality procedures are not required. The following ethical principles still "
    "apply.")
T.items([
    "**Lawful data use.** All data are publicly disclosed by the IDX and listed firms; any "
    "data obtained from commercial providers are used within their license terms and are not "
    "redistributed.",
    "**Integrity and transparency.** Hypotheses and windows are fixed in advance, all results "
    "are reported whether or not they support the hypotheses, and data and code are archived "
    "so the analysis can be verified.",
    "**Originality.** All sources are cited in APA 7th style, and the final thesis will be "
    "checked with Turnitin against the department’s maximum similarity of 25%. Any use of "
    "AI-assisted writing tools will be disclosed in line with department policy; the research "
    "idea, analysis and interpretation remain the author’s responsibility.",
    "**No investment advice.** Findings describe historical return patterns and will not be "
    "presented as recommendations to buy or sell specific stocks.",
], style="alpha")
T.h3("3.6.2 Thesis Timeline")
T.p("The thesis is planned to be completed within eight months after the proposal is approved, "
    "including the journal manuscript draft and the Three Minutes Thesis video required for "
    "graduation. Figure 3.2 and Table 3.9 present the schedule.")
T.figure(figdir / "fig_gantt.png", "3.2", "Thesis Completion Gantt Chart", width_cm=13.0,
         source="Source: Planned by the author.")
T.caption_table("3.9", "Thesis Completion Timeline")
T.table(
    ["No", "Activity", "Period", "Output"],
    [
        ["1", "Proposal revision and supervisor approval", "Dec 2026 – mid Jan 2027",
         "Approved proposal"],
        ["2", "Collect LQ45 lists, announcement dates and EPS", "Jan – Feb 2027",
         "Raw financial data set"],
        ["3", "Collect daily prices and IHSG; data cross-checks", "mid Jan – Feb 2027",
         "Raw price data set; error log"],
        ["4", "Construct SUE, SUER and CAR; data cleaning", "mid Feb – Mar 2027",
         "Analysis data set"],
        ["5", "Hypothesis tests and robustness checks", "mid Mar – Apr 2027",
         "Result tables and figures"],
        ["6", "Write Chapter IV", "mid Apr – May 2027", "Chapter IV draft"],
        ["7", "Write Chapter V; full draft review", "mid May – mid Jun 2027", "Full thesis draft"],
        ["8", "Journal manuscript draft and 3MT video", "Jun 2027", "Manuscript; video"],
        ["9", "Final revision, plagiarism check and defense", "mid Jun – Jul 2027",
         "Final thesis"],
    ],
    widths=[1.0, 5.6, 3.9, 3.5], size=10, align_center_cols=(0,))
T.source("Source: Planned by the author.")

# ------------------------------------------------------------------------ REFERENCES
T.page_break()
T.front_title("DAFTAR PUSTAKA")
refs = [
    "Ball, R., & Brown, P. (1968). An empirical evaluation of accounting income numbers. "
    "*Journal of Accounting Research, 6*(2), 159–178. https://doi.org/10.2307/2490232",
    "Ball, R., & Brown, P. (2019). Ball and Brown (1968) after fifty years. *Pacific-Basin "
    "Finance Journal, 53*, 410–431. "
    "https://www.sciencedirect.com/science/article/pii/S0927538X18306395",
    "Barberis, N., Shleifer, A., & Vishny, R. (1998). A model of investor sentiment. "
    "*Journal of Financial Economics, 49*(3), 307–343. "
    "https://doi.org/10.1016/S0304-405X(98)00027-0",
    "Bernard, V. L., & Thomas, J. K. (1989). Post-earnings-announcement drift: Delayed price "
    "response or risk premium? *Journal of Accounting Research, 27*, 1–36. "
    "https://doi.org/10.2307/2491062",
    "Bernard, V. L., & Thomas, J. K. (1990). Evidence that stock prices do not fully reflect "
    "the implications of current earnings for future earnings. *Journal of Accounting and "
    "Economics, 13*(4), 305–340. https://doi.org/10.1016/0165-4101(90)90008-R",
    "Brown, S. J., & Warner, J. B. (1985). Using daily stock returns: The case of event "
    "studies. *Journal of Financial Economics, 14*(1), 3–31. "
    "https://doi.org/10.1016/0304-405X(85)90042-X",
    "Cai, G., Lin, B., Wei, M., & Xu, X. (2021). The role of institutional investors in "
    "post-earnings announcement drift: Evidence from China. *Accounting and Business "
    "Research, 51*(2), 206–236. https://doi.org/10.1080/00014788.2020.1773755",
    "Chordia, T., Goyal, A., Sadka, G., Sadka, R., & Shivakumar, L. (2009). Liquidity and the "
    "post-earnings-announcement drift. *Financial Analysts Journal, 65*(4), 18–32. "
    "https://doi.org/10.2469/faj.v65.n4.3",
    "Dargenidou, C., Tonks, I., & Tsoligkas, F. (2018). Insider trading and the post-earnings "
    "announcement drift. *Journal of Business Finance & Accounting, 45*(3–4), 482–508. "
    "https://doi.org/10.1111/jbfa.12305",
    "DellaVigna, S., & Pollet, J. M. (2009). Investor inattention and Friday earnings "
    "announcements. *The Journal of Finance, 64*(2), 709–749. "
    "https://doi.org/10.1111/j.1540-6261.2009.01447.x",
    "Fama, E. F. (1970). Efficient capital markets: A review of theory and empirical work. "
    "*The Journal of Finance, 25*(2), 383–417. https://doi.org/10.2307/2325486",
    "Fama, E. F. (1998). Market efficiency, long-term returns, and behavioral finance. "
    "*Journal of Financial Economics, 49*(3), 283–306. "
    "https://doi.org/10.1016/S0304-405X(98)00026-9",
    "Fama, E. F., Fisher, L., Jensen, M. C., & Roll, R. (1969). The adjustment of stock prices "
    "to new information. *International Economic Review, 10*(1), 1–21. "
    "https://doi.org/10.2307/2525569",
    "Hirshleifer, D., Lim, S. S., & Teoh, S. H. (2009). Driven to distraction: Extraneous "
    "events and underreaction to earnings news. *The Journal of Finance, 64*(5), 2289–2325. "
    "https://doi.org/10.1111/j.1540-6261.2009.01501.x",
    "Ho, T. Q., Nguyen, Y., & Tran, H. (2024). The impact of insider ownership and "
    "institutional ownership on post-earnings-announcement-drift: Evidence from Vietnam. "
    "*Research in International Business and Finance, 70*, 102352. "
    "https://doi.org/10.1016/j.ribaf.2024.102352",
    "Hung, M., Li, X., & Wang, S. (2015). Post-earnings-announcement drift in global markets: "
    "Evidence from an information shock. *The Review of Financial Studies, 28*(4), "
    "1242–1283. https://doi.org/10.1093/rfs/hhu092",
    "Irawati. (2025, December 23). KSEI catat investor pasar modal tembus 20,12 juta [KSEI "
    "records capital market investors exceeding 20.12 million]. *Infobanknews*. "
    "https://infobanknews.com/ksei-catat-investor-pasar-modal-tembus-2012-juta/",
    "Fitch Ratings. (2026, March 4). *Fitch revises Indonesia’s outlook to negative; affirms "
    "at ‘BBB’* [Rating action commentary]. https://www.fitchratings.com/research/sovereigns/"
    "fitch-revises-indonesia-outlook-to-negative-affirms-at-bbb-04-03-2026",
    "Moody’s Ratings. (2026, February 5). *Rating action: Moody’s Ratings changes Indonesia’s "
    "outlook to negative from stable, affirms Baa2 rating* [Press release]. Bank Indonesia. "
    "https://www.bi.go.id/en/publikasi/ruang-media/news-release/Documents/"
    "SP_282726_rating-action-moodys.pdf",
    "Livnat, J., & Mendenhall, R. R. (2006). Comparing the post-earnings announcement drift "
    "for surprises calculated from analyst and time series forecasts. *Journal of Accounting "
    "Research, 44*(1), 177–205. https://doi.org/10.1111/j.1475-679X.2006.00196.x",
    "MSCI. (2026, October 15). *MSCI launches a consultation on a proposal to reclassify "
    "Indonesia from Emerging Market to Standalone Market status* [Press release]. "
    "https://www.msci.com/downloads/documents/press-releases/media-room/Indonesia_PR_Oct_26.pdf",
    "Mendenhall, R. R. (2004). Arbitrage risk and post-earnings-announcement drift. *The "
    "Journal of Business, 77*(4), 875–894. https://doi.org/10.1086/422627",
    "Nainggolan, R. (2025). Examining the information efficiency of Indonesia equity market: "
    "Evidence from earnings announcements. *International Journal of Economics and Financial "
    "Issues, 15*(1), 265–271. https://doi.org/10.32479/ijefi.16797",
    "Petersen, M. A. (2009). Estimating standard errors in finance panel data sets: Comparing "
    "approaches. *The Review of Financial Studies, 22*(1), 435–480. "
    "https://doi.org/10.1093/rfs/hhn053",
    "Piyawannasuth, P. (2025). *The information spillover effects of earnings announcements: "
    "Evidence from Thailand* [Independent study, Chulalongkorn University]. Chula Digital "
    "Collections. https://digital.car.chula.ac.th/chulaetd/74875",
    "Ramdhani, M. A. (2021). *Pengaruh komponen earnings surprise dan likuiditas saham "
    "terhadap post-earnings announcement drift (Studi di Bursa Efek Indonesia pada tahun "
    "2017–2019)* [The effect of earnings surprise components and stock liquidity on "
    "post-earnings announcement drift: A study on the Indonesia Stock Exchange, 2017–2019] "
    "[Undergraduate thesis, Universitas Gadjah Mada]. ETD UGM. "
    "http://etd.repository.ugm.ac.id/",
    "Sekaran, U., & Bougie, R. (2016). *Research methods for business: A skill-building "
    "approach* (7th ed.). Wiley.",
    "Shleifer, A., & Vishny, R. W. (1997). The limits of arbitrage. *The Journal of Finance, "
    "52*(1), 35–55. https://doi.org/10.1111/j.1540-6261.1997.tb03807.x",
    "Wijayana, S., & Achjari, D. (2020). Market reaction to the announcement of an information "
    "technology investment: Evidence from Indonesia. *Information & Management, 57*, "
    "103248. https://doi.org/10.1016/j.im.2019.103248",
]
for r in sorted(refs, key=lambda x: x.replace("*", "").replace("’", "'").lower()):
    T.reference(r)

T.save(out)
print("saved", out)
