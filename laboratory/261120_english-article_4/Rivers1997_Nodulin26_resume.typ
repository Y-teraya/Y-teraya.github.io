// Rivers et al. (1997) レジュメ（英論発表 4回目，A4 2ページ）
// 発表スライド「NOD26_論文紹介_発表用」と対応づけ．〔 〕内は対応するスライド
// コンパイル: typst compile Rivers1997_レジュメ.typ
#let green = rgb("#1b5e20")
#set page(paper: "a4", margin: (x: 2cm, y: 1.8cm), numbering: "1")
#set text(
  font: ("New Computer Modern", "Noto Serif JP"), weight: 450,
  size: 9.5pt,
  lang: "ja",
)
#set text(font: ("New Computer Modern", "Noto Serif JP"), weight: 450)
#show heading: set text(font: ("Inter 18pt", "Noto Sans JP"), weight: 450)
#show strong: set text(font: ("Inter 18pt", "Noto Sans JP"), weight: 300)
#show math.equation: set text(font: ("New Computer Modern Math", "Noto serif JP"))
#let lh = 1.1em
#set par(justify: true, leading: lh, spacing: lh + 0.15em)
#set list(indent: 0.4em, body-indent: 0.5em, spacing: lh)
#set heading(numbering: "1.1")
#show heading.where(level: 1): it => [
  #v(0.6em)
  #text(size: 11.5pt, weight: "bold", fill: green)[#it]
  #v(-0.5em)
  #line(length: 100%, stroke: 0.6pt + green)
  #v(0.2em)
]
#show heading.where(level: 2): it => [
  #v(0.3em)
  #text(size: 10pt, weight: "bold")[#it]
  #v(0.1em)
]
#let tbl(..args) = table(
  stroke: 0.4pt + luma(150),
  inset: 4pt,
  fill: (x, y) => if x == 0 { luma(235) } else { none },
  ..args,
)

// 対応スライドの表示
#let sl(body) = text(size: 8pt, weight: "regular", fill: luma(110))[〔#body〕]
// 記号
#let Pf = [$P_f$]
#let pf = [$p_f$]
#let Pd = [$P_d$]
#let Ps = [$P_"solute"$]
#let Ea = [$E_a$]
#let Hg = [HgCl#sub[2]]
// 主要な式（スライドの式3・式4）
#let eq3 = $ (dif V(t)) / (dif t) = P_f dot "SAV" dot "MVW" dot lr(\{ C_"in" / V(t) - C_"out" \}) $
#let eq4 = $ p_f = P_f / "SuD" $
#let eqR = $ P_f / P_d cases(= 1 & ": 膜に溶けて通る", > 1 & ": 孔の中を一列で通る") $

#align(right, [2026.11.20 寺谷優輝])
#block(inset: (left: -0.3em))[
  英語論文紹介 Journal name: _Journal of Boilogical Chemistry_ 272(26): 16256--16261.
]
#align(center)[
  #text(size: 14pt, weight: "bold", fill: green)[Functional Analysis of Nodulin 26, an Aquaporin in Soybean Root Nodule Symbiosomes]
  #v(0.1em)
]
#v(0.4em)

#tbl(
  columns: (2cm, 1fr),
  [著者・掲載誌], [Rivers RL, Dean RM, Chandy G, Hall JE, Roberts DM, Zeidel ML (1997) #emph[J Biol Chem] 272: 16256–16261.],
  [目的], [シンビオソーム膜（SM）の主要タンパク質ノジュリン26（NOD 26）が何を通すかを，卵母細胞（異種発現）と天然の SM 小胞で調べる．],
  [要点], [
    + NOD 26 は水銀感受性の#strong[水チャネル]．
    + 単位伝導度は #strong[AQP1 の約1/30]，水は孔を#strong[一列]で通る（#Pf / #Pd = 18.3）．
    + グリセロール・ホルムアミドも通すが尿素は通さず，溶質は#strong[水とは別の経路]を通ると推論．
  ],
)

= Introduction #sl[要旨]

- NOD 26 は SM の主要タンパク質で，アクアポリン（AQP）ファミリーに属する．C 末端 Ser-262 がリン酸化される．
- 平面脂質二重膜ではイオンチャネルになるが，似た AQP0 は卵母細胞で水チャネルとして働くと分かり，再評価されていた．
- #strong[問い]：NOD 26 は何を通すのか．#strong[天然の膜]でも測定する．

= Material & Method #sl[材料と方法]

#tbl(
  columns: (2.6cm, 1fr),
  [① 卵母細胞], [NOD 26 cRNA または水（対照）を注入．低張液で膨らむ速さから #Pf（式1・2）．3 mM #Hg で阻害，βME で回復．],
  [② SM 小胞], [大豆根粒からシンビオソームを単離し，CF を封入した SM 小胞を回収．直径は準弾性光散乱法．],
  [測定法], [#strong[ストップトフロー蛍光法]（デッドタイム 0.7 ms）．#Pf：外液の浸透圧を2倍 → 縮む（式3）．#Pd：ANTS 封入，重水と交換．#Ps：溶質を詰め，半分の濃度の液と混合（式5）．水銀は 0.1 mM #Hg．],
  [単位伝導度 #pf], [小胞の総表面積と NOD 26 量（SDS-PAGE・ウェスタン）からチャネル密度を出し，#Pf を割る（式4）．],
)

== 主要な式 #sl[式4]

#tbl(
  columns: (2.6cm, 1fr),
  [式4 #pf の算出], [#eq4 膜全体の #Pf を，単位面積あたりのチャネル数 SuD（個/cm#super[2]）で割り，#strong[チャネル1個あたり]の水の通しやすさ（cm#super[3]/s）にする．これにより AQP1 と1個単位で比較できる．],
  [#Pf / #Pd 比 一列輸送の判定], [#eqR 水が1分子ずつ独立に通るなら #Pf = #Pd．細い孔で水分子が#strong[一列]に並ぶと，浸透圧で押す流れは列ごと動くので #Pf の方が大きくなる．比は孔に並ぶ水分子の数の目安で，NOD 26 は18.3（AQP1 は13）．],
)

= Results #sl[結果]

#tbl(
  columns: (2.2cm, 1fr, 3.6cm),
  [#strong[図1] 卵母細胞], [NOD 26 で #Pf が有意に上昇（#emph[p] < 0.001）．#Hg で対照レベルへ低下し，βME で回復．], [水銀感受性の水チャネル],
  [#strong[図2] 小胞径], [山が一つ．直径 217・240 nm → 平均半径 228 nm．], [計算に使える単一集団],
  [#strong[図3] #Pf], [#strong[$0.05 plus.minus 0.003 #h(0.25em) upright(c m\/s)$]（20 °C，$n = 15$）．#Hg で 0.0038 cm/s へ劇的に低下．], [透過性が高く，水銀感受性],
  [#strong[図4] #Ea], [#strong[$3.3 ± 0.4 #h(0.25em) upright(k c a l\/m o l)$]（アレニウスプロット，12〜39 °C）．], [低い → チャネルを通る],
  [#strong[式4] #pf], [NOD 26 は SM 総タンパク質の約10%．#pf = #strong[$3.2 × 10^(−15) #h(0.25em) upright(c m^3\/s)$]．], [AQP1 の約$1/30$（中程度）],
  [#strong[図5] #Pd], [#Pd = 0.0024 cm/s（12 °C）．同温 #Pf = 0.044 → #Pf / #Pd = #strong[$18.3$]（AQP1 は13）．], [水は一列で通る],
  [#strong[図6・7] #Ps], [ホルムアミド > グリセロール ≫ アセトアミド > 尿素．前二者は #Hg で 85〜90% 低下，尿素は阻害されない．], [一部の溶質も通す],
)

= Discussion #sl[考察]

== A. シンビオソーム膜での役割 #sl[考察 A-1〜A-2]

- #strong[水輸送]：SM の水輸送は，高速・水銀感受性・低 #Ea ・一列という AQP の特徴をすべて示した．根粒は干ばつや塩分に敏感なので，NOD 26 は#strong[浸透圧調節]を担う可能性がある．
- #strong[溶質も NOD 26 が通す根拠]：
  + AQP にはグリセロールを通す仲間がいる．
  + 脂質膜ならグリセロールと尿素は同程度に通る．
  + 水銀はグリセロール・ホルムアミドだけを止める．
- #strong[残る疑問]：代謝物（リンゴ酸など）の輸送に関わるか．Ser-262 のリン酸化は輸送にどう影響するか．

== B. アクアポリンとしての NOD 26 #sl[考察 B-1〜B-2]

- #strong[選択性]：最も大きいグリセロール（51.4 cm#super[3]/mol）が尿素（32.8）より速い → 大きさだけでは決まらない．
- #strong[孔の長さ]：#Pf / #Pd 比は NOD 26 が18，AQP1 が13 → NOD 26 の一列の孔は長く，それが低い伝導度につながる可能性．
- #strong[別経路の推論]：水は一列で通るのに，水分子2個分より大きい溶質が通る → 溶質は水と別の経路を通ると推論した．

== 限界と今後の課題 #sl[考察 B-3]

- 精製 NOD 26 をプロテオリポソームに#strong[再構成]して確かめる必要がある．
- 天然の膜でのデータは，再構成系・異種発現系の結果を解釈する基準になる．

#pagebreak()

#set figure(gap: 0.4em)
#show figure.caption: set text(size: 8.5pt)
#show figure.caption: set align(left)

// 図を1枚置く関数：path = 画像ファイル，h = 画像枠の高さ，cap = キャプション
#let fig(path, h, cap) = figure(
  image(path, width: 100%, height: h, fit: "contain"),
  caption: cap,
)

#heading(numbering: none)[Figures]

#v(1fr)

#grid(
  columns: (1fr, 1fr),
  column-gutter: 0.8cm,
  row-gutter: 1cm,
  fig("fig1.png", 8.5cm)[
    #strong[卵母細胞の水透過性．]NOD 26 cRNA の注入で $P_f$ が対照より上昇し，HgCl#sub[2] で低下，βME で回復する（平均 ± S.D.，n = 4〜5）．
  ],
  fig("fig2.png", 8.5cm)[
    #strong[SM 小胞の直径分布．]準弾性光散乱法．山が一つの単一集団で，平均半径 228 nm．
  ],
  fig("fig3.png", 8.5cm)[
    #strong[SM 小胞の浸透水透過性（20 °C）．]外液の浸透圧を2倍にしたときの縮み方．0.1 mM HgCl#sub[2] で大きく遅れる（$P_f$ 0.05 → 0.0038 cm/s）．
  ],
  fig("fig4.png", 8.5cm)[
    #strong[アレニウスプロット．]12〜39 °C の ln #emph[k] と 1/#emph[T]．傾きから $E_a$ = 3.3 kcal/mol と低い．
  ],
)

#v(1fr)

#pagebreak()

#v(1fr)

// ---- 2ページ目：図5・図7（横並び）＋ 図6（全幅）----
#grid(
  columns: (1fr, 1fr),
  column-gutter: 0.8cm,
  fig("fig5.png", 8cm)[
    #strong[拡散水透過性（12 °C）．]H#sub[2]O を D#sub[2]O に置き換えると ANTS 蛍光が上昇．$P_d$ = 0.0024 cm/s，$P_f \/ P_d$ = 18.3．
  ],
  fig("fig7.png", 8cm)[
    #strong[溶質透過係数．]ホルムアミド・グリセロールは高く HgCl#sub[2] で 85〜90% 低下．尿素は阻害されない（平均 ± S.E.）．
  ],
)

#v(1fr)

#fig("fig6.png", 10.5cm)[
  #strong[溶質の流出に伴う小胞の縮み（20 °C）．]ホルムアミド・グリセロール・アセトアミド・尿素，HgCl#sub[2] なし／あり．溶質ごとに時間軸が異なる．
]

#v(1fr)