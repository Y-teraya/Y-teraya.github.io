// Doyle et al. (1997) レジュメ（卒論の背景部分の整理用）
// コンパイル: typst compile Doyle1997_レジュメ.typ
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
#set par(justify: true, leading: 0.65em, spacing: 0.6em)
#set list(indent: 0.4em, body-indent: 0.5em, spacing: 0.45em)
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

#align(right, [2026.10.16 寺谷優輝])
#block(inset: (left: -0.3em))[
  英語論文紹介 Journal name: _American Journal of Botany_ 84(4): 541--554.
]

#v(1.5em)

#align(center)[
  #text(size: 14pt, weight: "bold", fill: green)[A PHYLOGENY OF THE CHLOROPLAST GENE RBCL IN THE LEGUMINOSAE: TAXONOMIC CORRELATIONS AND INSIGHTS INTO THE EVOLUTION OF NODULATION]
  #v(0.1em)
  #text(size: 10pt)[― Doyle et al. (1997) ―]
]
#v(0.4em)

#tbl(
  columns: (2cm, 1fr),
  [論文], [Doyle JJ, Doyle JL, Ballenger JA, Dickson EE, Kajita T, Ohashi H (1997) #emph[Am J Bot] 84: 541–554.],
  [目的], [マメ科全体の #emph[rbcL] 系統解析により，亜科・連の分類と根粒形成の進化を検討する．],
  [要点], [① ジャケツイバラ亜科は非単系統である．② 無根粒が祖先的で，根粒形成は並行的に複数回獲得されたと最節約的に推定された．③ 根粒形成は，少数の遺伝子変異で失われうる複雑な形質である．],
)

= Introduction

== 歴史的な流れ

#tbl(
  columns: (2.2cm, 3.6cm, 1fr),
  [1865], [Bentham], [分類の基礎．以後の研究の土台となる．],
  [1981], [Polhill and Raven ほか], [3亜科の体系（Cronquist は3科説）．ジャケツイバラ亜科，Caesalpinieae，Sophoreae，Phaseoleae などは非単系統の疑いがあった．],
  [1990年代], [cpDNA 制限酵素地図，ITS，フィトクロム], [分子形質で系統仮説を構築．ただし単一の連，または近縁な数連が対象であった．],
  [1994–95], [Tucker and Douglas; Chappill], [科全体を扱う形態解析．結果が大きく異なった．],
  [1995], [Käss and Wink], [#emph[rbcL] による解析．Genisteae 連が中心であった．],
  [1997], [本論文], [科全体を対象とした #emph[rbcL] 系統解析．],
)

== 根粒形成についての遺伝的な議論

- #strong[複雑な現象]である：新規の器官（根粒），新規の細胞小器官（シンビオソーム），根粒関連タンパク質（ノデュリン）を伴う．ノデュリンの一部は新規遺伝子の産物の可能性がある．
- #strong[失われやすい]：単一遺伝子の変異で根粒形成を阻止できる．また，高コストで，窒素が十分な環境では起こらない．喪失は頻繁に生じうる．
- #strong[複数回の起源]：他の双子葉植物でも窒素固定共生は独立に生じた．マメ科の根粒は，解剖学的にも別起源が支持される．
- #strong[根粒の型は宿主が支配し，系統と対応する]：不定型は caesalpinioid 型，crotalarioid 型，lupinoid 型．定型は desmodioid 型，mucunoid 型，aeschynomenoid 型．Phaseoleae は定型で，窒素をウレイドとして輸送する．
- #strong[検証の方法]：ノデュリン遺伝子ファミリーのオルソロジーと発現パターンの解析が提案されている．

= Material & Method

#tbl(
  columns: (2cm, 1fr),
  [材料], [3亜科の属の個体（植物園・野外の個体，標本庫に登録）．栽培・処理は行っていない．ジャケツイバラ亜科は5連すべて，マメ亜科は30連中26連．基部（Caesalpinieae，Swartzieae，Sophoreae）を重点的にサンプリング．],
  [外群], [Polygalaceae（#emph[Securidaca]，#emph[Polygala]）など．],
  [配列取得], [全ゲノム DNA から #emph[rbcL] を PCR 増幅（難しい試料は booster PCR）．手動または自動シーケンサーで決定．一部の配列は文献・GenBank から追加．],
  [解析], [最初の1428 bp．重み付けなしの最節約法（PAUP，NONA）．主解析は84分類群．],
  [支持], [パーシモニー・ジャックナイフ法．],
)

= Results

== 系統解析の全体像

- 最節約樹は#strong[184本]（樹長1913，ci 0.35，ri 0.52）．
- マメ亜科基部（Sophoreae，Swartzieae）は多分岐で，支持が低い．派生的なマメ亜科は置換が多く，よく解像された．
- マメ科は常に単系統．Polygalaceae（単独，または Surianaceae とともに）が姉妹群となった．

== 分類群の自然性

#tbl(
  columns: (4cm, 1fr),
  [マメ亜科], [単系統（支持は弱い）．],
  [ネムノキ亜科], [単系統．Caesalpinieae 内に入れ子となった．],
  [ジャケツイバラ亜科], [#strong[非単系統]．樹の4か所に分散した．最初期の分岐は Cercideae または Dialiinae，次に Detarieae/Macrolobieae．],
  [Cassieae / Caesalpinieae], [Cassieae は多系統，Caesalpinieae は側系統．],
  [Sophoreae / Swartzieae], [基部で混在し，非単系統．50-kb 逆位をもつ属ともたない属が混在．],
  [Phaseoleae / Millettieae], [各亜連・属が大クレード内に分散．Desmodieae や Psoraleeae が Phaseoleae 内に入れ子となった．],
)

== 根粒形成の進化

- 近縁科や基部の群（Cercideae，Dialiinae，Detarieae/Macrolobieae）は根粒をつくらない．→ #strong[無根粒が祖先的]．
- 最節約的には#strong[3回の並行的獲得]：マメ亜科，ネムノキ亜科につながる祖先，#emph[Chamaecrista]．
- ただし支持の弱い関係が多く，ジャケツイバラ亜科での単一起源も否定できない．
- マメ亜科基部では根粒の有無が混在し，単一起源の位置は特定できない．喪失の可能性もある（#emph[Luetzelburgia]）．
- 根粒の型の分布は，#emph[rbcL] 樹のトポロジーとおおむね対応した．

= Discussion

== 本研究で明らかになったこと

- #strong[分類の自然性]：3亜科のうち自然群はマメ亜科とネムノキ亜科のみで，ジャケツイバラ亜科は不自然であった．連のレベルでも，Cassieae，Caesalpinieae，Sophoreae，Swartzieae，Phaseoleae，Millettieae が単系統とならず，Introduction で述べた非単系統の疑いが支持された．
- #strong[既存の仮説との関係]：#emph[rbcL] 樹は，Chappill (1995) の解析とよく一致し，Tucker and Douglas (1994) とは一致の程度が低かった．Bentham 以来の研究は，分子系統学の確かな基礎になっているといえる．
- #strong[形質の進化]：単性花をもつ #emph[Gymnocladus] と #emph[Ceratonia] が基部近くで姉妹群となり，花の退化は共有派生形質である可能性が示された．IR の喪失は，#emph[rbcL] 樹でも単一のクレードを形成したが，#emph[rbcL] と IR は同じ葉緑体ゲノムの一部なので，独立な検定にはならない．

== 根粒形成についての考察

- #strong[起源]：無根粒が祖先的で，最節約的には3回の並行的獲得となった．しかし，支持が弱く，Cassiinae の側系統性や近縁性も考えると，ジャケツイバラ亜科での単一起源も不合理ではない．
- #strong[マメ亜科基部]：根粒の有無が混在し，起源の位置は特定できない．祖先が根粒形成能をもち，一部の分類群が失った可能性もある．これは，単一遺伝子の変異で阻止でき，高コストで失われやすいという遺伝的な議論と整合する．
- #strong[複数回の獲得の妥当性]：複雑な共生系が複数回生じたことは，他の双子葉植物で独立に生じた例があること，根粒の解剖学的特徴が別起源を支持することから，あり得ると考えられる．
- #strong[根粒の型]：caesalpinioid 型は，ジャケツイバラ亜科，ネムノキ亜科，マメ亜科で共通し，大まかな構造では根粒が科内で相同に見える．型の分布は #emph[rbcL] 樹とおおむね対応した．

== 限界と今後の課題

- #strong[rbcL の限界]：基部の関係や，科の初期の放散については，解像も支持も不十分であった．#strong[他の遺伝子，とくに独立な核遺伝子]のデータが必要である．
- #strong[根粒の相同性]：ノデュリン遺伝子ファミリーのオルソロジーと発現パターンの解析で検証すべきである．
- #strong[未調査の群]：根粒形成の状態や窒素の輸送様式（ウレイドなど）が不明な分類群が多い．Indigofereae や Psoraleeae などは，調べる価値がある．

#pagebreak()

#v(1fr)

#figure(image("Fig1.png", width: 95%))

#v(1fr)

#pagebreak()

#v(1fr)

#figure(image("Fig2-5.png", width: 95%))

#v(1fr)

#figure(image("Fig6-7.png", width: 95%))

#v(1fr)