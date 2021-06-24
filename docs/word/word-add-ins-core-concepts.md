---
title: Office アドインの Word JavaScript オブジェクト モデル
description: Word 固有の JavaScript オブジェクト モデルの最も重要なクラスについて説明します。
ms.date: 10/14/2020
localization_priority: Priority
ms.openlocfilehash: 43ca88e7899e2ff11748dc91d5c8a5059d8bb559
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077233"
---
# <a name="word-javascript-object-model-in-office-add-ins"></a>Office アドインの Word JavaScript オブジェクト モデル

この記事では、[Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) を使用してアドインを構築するための基本的な概念について説明します。API を使用するための基本的なコア コンセプトを紹介します。

> [!IMPORTANT]
> Word API の非同期性と、ドキュメントでの動作方法については、「[アプリケーション固有の API モデルの使用](../develop/application-specific-api-model.md)」を参照してください。

## <a name="officejs-apis-for-word"></a>Word 用の Office.js API

Word アドインは、次の 2 つの JavaScript オブジェクト モデルを含む Office JavaScript API を使用して、Excel のオブジェクトを操作します:

* **Word JavaScript API**: [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) には、ドキュメント、範囲、テーブル、リスト、フォーマットなどにアクセスするために使用できる厳密に型指定されたオブジェクトが用意されています。

* **共通 API**: [共通 API](/javascript/api/office) を使用して、UI、ダイアログ、クライアント設定など、複数のタイプの Office アプリケーションに共通の機能にアクセスできます。

Word を対象にしたアドインでは、機能の大部分を Word JavaScript API を使用して開発する可能性がありますが、共通 API のオブジェクトも使用します。次に例を示します。

* [コンテキスト](/javascript/api/office/office.context): `Context` オブジェクトは、アドインのランタイム環境を表し、API の主要なオブジェクトへのアクセスを提供します。 これは `contentLanguage` や `officeTheme` などのドキュメント構成の詳細で構成され、`host` や `platform` などのアドインのランタイム環境に関する情報も提供します。 さらに、`requirements.isSetSupported()` メソッドも提供されます。これを使用すると、指定した要件セットが、アドインが実行されている Excel アプリケーションでサポートされているかどうかを確認できます。
* [ドキュメント](/javascript/api/office/office.document): `Document` オブジェクトは `getFileAsync()` メソッドを提供します。これを使用すると、アドインが実行されている Word ファイルをダウンロードできます。

![Word JS API と共通 API の違い。](../images/word-js-api-common-api.png)

## <a name="word-specific-object-model"></a>Word 固有のオブジェクト モデル

Word API について理解するには、ドキュメントの構成要素が互いにどのように関連しているかを理解する必要があります。

* **ドキュメント** には **セクション** と、設定やカスタム XML パーツなどのドキュメントレベルのエンティティが含まれます。
* **セクション** には **本文** が含まれます。
* **本文** は、特に **パラグラフ**、**ContentControl**、および **範囲** オブジェクトへのアクセスを提供します。
* **範囲** は、テキスト、空白、**テーブル**、画像など、コンテンツの連続した領域を表します。 また、テキストの操作方法のほとんどが含まれます。
* **リスト** は、番号付きまたは箇条書きのリスト内のテキストを表します。

## <a name="see-also"></a>関連項目

- [Word JavaScript API の概要](../reference/overview/word-add-ins-reference-overview.md)
- [最初の Word アドインをビルドする](../quickstarts/word-quickstart.md)
- [Word アドインのチュートリアル](../tutorials/word-tutorial.md)
- [Word JavaScript API リファレンス](/javascript/api/word)
- [Microsoft 365 開発者プログラムについて](https://developer.microsoft.com/microsoft-365/dev-program)
