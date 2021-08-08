---
title: Office アドインの文字体裁ガイドライン
description: アドインで使用する書体とフォント サイズOfficeします。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 8cc17a25ed33fc34dd7a44622baacc620304402931de87eeadee903db5f135b0
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57082311"
---
# <a name="typography"></a>文字体裁

Segoe は、Office の標準的な書体です。 Office の作業ウィンドウ、ダイアログ ボックス、コンテンツ オブジェクトと調和するように、アドインで使用します。 [Fabric Core を](fabric-core.md) 使用すると、Segoe にアクセスできます。 フォントの太さからサイズまで数多くのバリエーションで Segoe の完全な文字体裁を、便利な CSS クラスで提供します。 一部の Fabric Core のサイズと重み付けは、一部のアドインOffice見える場合があります。 調和的に収まるか、競合を回避するには、Fabric Core タイプ のランプのサブセットの使用を検討してください。 次の表に、アドインで使用することをお勧めする Fabric Core の基本Office示します。

> [!NOTE]
> これらの基底クラスには、テキストの色が含まれていません。 白い背景のほとんどのテキストには、Fabric Core の "ニュートラル プライマリ" を使用します。
>
> 使用可能なタイポグラフィの詳細については [、「Web Typography」を参照してください](https://developer.microsoft.com/fluentui#/styles/web/typography)。

|型 |クラス |サイズ |太さ |お勧めの用法 |
|------ |----- |---- |------ |----------------- |
|ヒーロー|.ms-font-xxl |28 px | Segoe Light |<ul><li>これは、Office の他のすべての文字体裁の要素よりも大きいクラスです。視覚的な階層から外れないように、慎重に使用します。</li><li>制約のある領域内で長い文字列に対して使用しないでください。</li><li>このクラスを使用して、テキストの周りに十分な余白を確保してください。</li><li>通常、最初の実行メッセージ、ヒーロー要素、その他の行動喚起に使用します。</li></ul> |
|タイトル|.ms-font-xl |21 px |Segoe Light | <ul><li>このクラスは、Office アプリケーションの作業ウィンドウ タイトルと一致します。</li><li>文字体裁の階層が平板にならないように、慎重に使用します。</li><li>通常、ダイアログ ボックス、ページ、コンテンツ タイトルなどの最上位の要素として使用します。</li></ul> |
|サブタイトル|.ms-font-l |17 px |Segoe Semilight | <ul><li>このクラスは、タイトルの 1 つ下です。</li><li>通常、サブタイトル、ナビゲーション要素、グループ ヘッダーとして使用します。</li><ul> |
|Body|.ms-font-m |14 px |Segoe Regular |<ul><li>通常、アドイン内の本文として使用します。</li><ul>|
|Caption|.ms-font-xs |11 px | Segoe Regular |<ul><li>通常、タイムスタンプなどの 2 番目や 3 番目のテキストとして、行、キャプション、フィールド ラベルごとに使用します。</li><ul>|
|Annotation|.ms-font-mi |10 px |Segoe Semibold |<ul><li>文字体裁の最小の階層は、稀にしか使用しないでください。読みやすさが求められない状況で使用できます。</li><ul>|
