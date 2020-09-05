---
title: Office アドインの文字体裁ガイドライン
description: Office アドインで使用する書体とフォントサイズについて説明します。
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: d7347e2e6ee01386d631fea8c2b388ad5b61005e
ms.sourcegitcommit: 10463841a977e9b8415362a3ae91b0ae5eebbf89
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/04/2020
ms.locfileid: "47399565"
---
# <a name="typography"></a>文字体裁

Segoe は、Office の標準的な書体です。 Office の作業ウィンドウ、ダイアログ ボックス、コンテンツ オブジェクトと調和するように、アドインで使用します。 Office UI Fabric では、Segoe にアクセスできます。 フォントの太さからサイズまで数多くのバリエーションで Segoe の完全な文字体裁を、便利な CSS クラスで提供します。 すべての Office UI Fabric のサイズと太さが Office アドインで適切に表示されるわけではありません。 調和よく収めるため、または競合を回避するために、Fabric 文字体裁のサブセットを使うことを検討してください。 次の表は、Office アドインで使用することを推奨する Fabric の基本クラスを示しています。

> [!NOTE]
> これらの基底クラスには、テキストの色が含まれていません。 背景が白の場合、ほとんどのテキストには Fabric の "中間色" を使用します。
>
> 利用可能な文字体裁の詳細については、「 [Web タイポグラフィ](https://developer.microsoft.com/fluentui#/styles/web/typography)」を参照してください。

|Type |クラス |サイズ |太さ |お勧めの用法 |
|------ |----- |---- |------ |----------------- |
|ヒーロー|.ms-font-xxl |28 px | Segoe Light |<ul><li>これは、Office の他のすべての文字体裁の要素よりも大きいクラスです。視覚的な階層から外れないように、慎重に使用します。</li><li>制約のある領域内で長い文字列に対して使用しないでください。</li><li>このクラスを使用して、テキストの周りに十分な余白を確保してください。</li><li>通常、最初の実行メッセージ、ヒーロー要素、その他の行動喚起に使用します。</li></ul> |
|Title|.ms-font-xl |21 px |Segoe Light | <ul><li>このクラスは、Office アプリケーションの作業ウィンドウ タイトルと一致します。</li><li>文字体裁の階層が平板にならないように、慎重に使用します。</li><li>通常、ダイアログ ボックス、ページ、コンテンツ タイトルなどの最上位の要素として使用します。</li></ul> |
|サブタイトル|.ms-font-l |17 px |Segoe Semilight | <ul><li>このクラスは、タイトルの 1 つ下です。</li><li>通常、サブタイトル、ナビゲーション要素、グループ ヘッダーとして使用します。</li><ul> |
|Body|.ms-font-m |14 px |Segoe Regular |<ul><li>通常、アドイン内の本文として使用します。</li><ul>|
|Caption|.ms-font-xs |11 px | Segoe Regular |<ul><li>通常、タイムスタンプなどの 2 番目や 3 番目のテキストとして、行、キャプション、フィールド ラベルごとに使用します。</li><ul>|
|Annotation|.ms-font-mi |10 px |Segoe Semibold |<ul><li>文字体裁の最小の階層は、稀にしか使用しないでください。読みやすさが求められない状況で使用できます。</li><ul>|
