---
title: Office アドインの文字体裁ガイドライン
description: ''
ms.date: 06/27/2018
ms.openlocfilehash: b9c5a957411a7c2df078be54df514237280cd150
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432957"
---
# <a name="typography"></a>文字体裁

Segoe は、Office の標準的な書体です。Office の作業ウィンドウ、ダイアログ ボックス、コンテンツ オブジェクトと調和するように、アドインで使用します。Office UI Fabric では、Segoe にアクセスできます。フォントの太さからサイズまで数多くのバリエーションで Segoe の完全な文字体裁を、便利な CSS クラスで提供します。すべての Office UI Fabric のサイズと太さが Office アドインで適切に表示されるわけではありません。調和よく収めるため、または競合を回避するために、Fabric 文字体裁のサブセットを使うことを検討してください。Office アドインでの使用をお勧めする Fabric の基底クラスの一覧を次に示します。

|サンプル |クラス |サイズ |太さ |お勧めの用法 |
|------ |----- |---- |------ |----------------- |
|![ヒーロー テキスト イメージ](../images/add-in-typeramp-hero.png)|.ms-font-xxl |28 px | Segoe Light |<ul><li>これは、Office の他のすべての文字体裁の要素よりも大きいクラスです。視覚的な階層から外れないように、慎重に使用します。</li><li>制約のある領域内で長い文字列に対して使用しないでください。</li><li>このクラスを使用して、テキストの周りに十分な余白を確保してください。</li><li>通常、最初の実行メッセージ、ヒーロー要素、その他の行動喚起に使用します。</li></ul> |
|![ヒーロー テキスト イメージ](../images/add-in-typeramp-title.png)|.ms-font-xl |21 px |Segoe Light | <ul><li>このクラスは、Office アプリケーションの作業ウィンドウ タイトルと一致します。</li><li>文字体裁の階層が平板にならないように、慎重に使用します。</li><li>通常、ダイアログ ボックス、ページ、コンテンツ タイトルなどの最上位の要素として使用します。</li></ul> |
|![ヒーロー テキスト イメージ](../images/add-in-typeramp-subtitle.png)|.ms-font-l |17 px |Segoe Semilight | <ul><li>このクラスは、タイトルの 1 つ下です。</li><li>通常、サブタイトル、ナビゲーション要素、グループ ヘッダーとして使用します。</li><ul> |
|![ヒーロー テキスト イメージ](../images/add-in-typeramp-body.png)|.ms-font-m |14 px |Segoe Regular |<ul><li>通常、アドイン内の本文として使用します。</li><ul>|
|![ヒーロー テキスト イメージ](../images/add-in-typeramp-caption.png)|.ms-font-xs |11 px | Segoe Regular |<ul><li>通常、タイムスタンプなどの 2 番目や 3 番目のテキストとして、行、キャプション、フィールド ラベルごとに使用します。</li><ul>|
|![ヒーロー テキスト イメージ](../images/add-in-typeramp-annotation.png)|.ms-font-mi |10 px |Segoe Semibold |<ul><li>文字体裁の最小の階層は、稀にしか使用しないでください。読みやすさが求められない状況で使用できます。</li><ul>|

> [!NOTE]
> これらの基底クラスには、テキストの色が含まれていません。背景が白の場合、ほとんどのテキストには Fabric の "中間色" を使用します。
