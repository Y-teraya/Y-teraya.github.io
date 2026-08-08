#import "@preview/rubber-article:0.5.2": *

#set text(font: ("New Computer Modern", "Noto Serif JP"), weight: 450)
#show heading: set text(font: ("Inter 18pt", "Noto Sans JP"), weight: 450)
#show strong: set text(font: ("Inter 18pt", "Noto Sans JP"), weight: 300)
#set par(first-line-indent: 1em, spacing: 0.5em, leading: 0.5em)
#set heading(numbering: "1.")
#show heading: it => {
    it
    par(text(size: 0em, ""))
}

#set page(
  paper: "a4",
  numbering: "1",
)

#align(right, [2026.07.09. 寺谷優輝])
#block(inset: (left: -0.3em))[
  英語論文紹介02 Journal Name: _PLoS Computational Biology_, Vol: 6, Pages: 1 – 8.
]

#v(-3.5em)

#maketitle(
  title: "Computational Complementation: A Modelling Approach to Study Signalling Mechanisms during Legume Autoregulation of Nodulation",
  authors: (
    "Liqi Han, Jim Hanan & Peter M. Gresshoff",
  ),
)

#v(-3.5em)

= 背景と目的

== 背景

マメ科植物と根粒菌の共生による窒素固定は，世界の主要作物生産量の27%，加工植物油の35%以上を支える重要なプロセスである．しかし宿主植物にとって過剰な根粒形成は代謝資源の浪費や側根形成への悪影響を招くため，植物は「根粒形成の自己調節（AON）」と呼ばれる長距離シグナル伝達系によって根粒数を制御している．

== 既報の知見

根粒原基の誘導により移動性シグナルQが生成され，木部を通って葉に運ばれる．葉の維管束にあるLRR型受容体キナーゼ（ダイズではGmNARK，ロータスではHAR1，メディカゴではSUNN）がQを感知し，地上部由来阻害物質（SDI）の産生を誘導，これが根に運ばれて根粒形成をさらに阻害すると仮説付けられている．SDIは野生型葉から抽出・再投与することで機能欠損変異体（nts1116など）の表現型を野生型に回復させることが確認されている．

== 未解明の課題

AONのフラックスデータや生化学的データが乏しく，NARK前後のシグナル伝達の詳細（産生・輸送・知覚・機能）の多くは未解明である．特に，子葉がAONにおいてSDI産生器官として機能するか（子葉-地上部仮説）、それとも根の一部として機能しSDIを産生しないか（子葉-根仮説）は不明であった．

== 本研究の目的

実験的検証が困難なAON仮説を扱うため，「計算補完（Computational Complementation）」という新しい機能構造モデリング手法を開発し，その最初の応用例として子葉のSDI産生関与を検証する．

= 材料および方法

== 供試材料

- ダイズ（野生型 Bragg および根粒過形成変異体 nts1116，_Glycine max_ L. Merrill）
- 根粒菌（_Bradyrhizobium japonicum_ CB1809株）

== モデル構築手法

L-systemに基づく言語「cpfg」（L-studio）を用い，以下の手順でモデルを構築した：
+ Bragg・nts1116それぞれの生体計測データに基づくアーキテクチャモデルの構築
+ nts1116アーキテクチャモデルを，シグナル伝達（産生・輸送・知覚・機能）を統合した機能構造モデルへ拡張
+ 確認済み・仮説上のAONメカニズムによるパラメータ化（"nts1116+AON"モデル）
+ 生成された表現型とBragg実測パターンとの比較
+ 支持された仮説の実植物実験による検証

#figure(image("./fig0-1.png", height: 15%), caption: [根粒分布の生体計測データ測定])

== 仮想実験デザイン

子葉-根仮説（CRH_1〜CRH_27）と子葉-地上部仮説（CSH_1〜CSH_27）それぞれについて，QおよびSDIの輸送速度（60・160・360 mm/日）と産生・阻害閾値比 $Q_"ini" : "SDI"_"ihbt"$（0.5・1・2）を組み合わせた27条件の仮想実験を実施した．

== 実植物接ぎ木実験

Braggとnts1116の間で，子葉の有無を入れ替えた4種類の接ぎ木（Ns+Nc+Br，Ns+Bc+Br，Bs+Bc+Nr，Bs+Nc+Nr）を作成し，接ぎ木後2週間の根粒数を子葉保持状態別に計数した．

= 測定項目および分析・統計方法

== 類似度指標
複相補結果とBragg表現型の一致度を次式で定量化した：
$ S_(c p) = (N_(n t) - N_(c p)) / (N_(n t) - N_(b r)) $
（$N_(n t)$: nts1116根粒数，$N_(b r)$: Bragg根粒数，$N_(c p)$: 補完モデル根粒数）

== 評価基準
$S_(c p)$が80〜120%の範囲にある場合を「良好」，それ以外を「不良」とした．

== データ収集
播種後10日目・16日目における全根粒数の類似度，および播種後16日目の根粒分布パターンを比較した．

= 結果

== Figure 3・4
播種後10日目では子葉-根仮説の全実験が「不良」であった一方，子葉-地上部仮説では多数の「良好」な結果が得られた．16日目には子葉-根仮説で4件（CRH_1, 2, 11, 13），子葉-地上部仮説で12件が「良好」となり，一貫して子葉-地上部仮説が優勢であった．

== Figure 6
CRH_1等で根粒数は良好な類似度を示したものの，根粒サイズ・密度の分布はBraggパターンから乖離していた．対照的にCSH_1が生成する根粒分布はBraggの構造モデルに近似していた．

== Figure 7
実接ぎ木実験の結果，Bragg由来の子葉を持つ接ぎ木体（Ns+Nc+Br，Bs+Bc+Nr）は，子葉を持たない対応個体より根粒数が多く，子葉保持状態（脱落・黄変・緑色）によっても根粒数に差が見られ，子葉がSDI供給に関与することが確認された．

== Figure 8
仮想実験モデルに立ち返った可視化により，CRH_1のSDI濃度は生育初期にCSH_1より低いが10日目以降に逆転して高くなり，実接ぎ木実験で観察された非線形的な根粒形成差異と一致した．

= 考察

== 結論
+ 「計算補完」という新しい機能構造モデリング手法の考案
+ 子葉が地上部の一部としてSDIを産生するという仮説の仮想実験による支持
+ 実植物接ぎ木実験によるその予測の確認

#h(1fr)の3点が新たに明らかになった．

== 意義・展望
本手法は，実験的検証が困難な内部シグナル伝達仮説を，観測可能な植物構造をレポーターとして評価することを可能にする．今後はCLEペプチド（Q候補）やオーキシン（SDI候補）の検証，土壌窒素状態などの環境要因の影響評価，および母株の_Bradyrhizobium_感染状態を介した世代間効果の検証への応用が期待される．

#pagebreak()