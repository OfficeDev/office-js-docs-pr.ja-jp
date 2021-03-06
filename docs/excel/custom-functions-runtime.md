---
ms.date: 09/25/2020
description: 作業Excel特定の JavaScript ランタイムを使用しないカスタム関数について説明します。
title: UI レス のカスタム関数Excelランタイム
localization_priority: Normal
ms.openlocfilehash: aa2cf2632ddf9eb1ad1eb202b031ee2ca686af01
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349624"
---
# <a name="runtime-for-ui-less-excel-custom-functions"></a>UI レス のカスタム関数Excelランタイム

作業ウィンドウを使用しないカスタム関数 (UI レスのカスタム関数) は、計算のパフォーマンスを最適化するように設計された JavaScript ランタイムを使用します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

この JavaScript ランタイムは、UI レスのカスタム関数と作業ウィンドウでデータを格納するために使用できる名前空間内の `OfficeRuntime` API へのアクセスを提供します。

## <a name="requesting-external-data"></a>外部データの要求

UI レスのカスタム関数内では [、Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) のような API を使用するか、サーバーとやり取りするための HTTP 要求を発行する標準 Web API [である XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest)を使用して外部データを要求できます。

XMLHttpRequests を作成する場合は、UI レス関数で追加のセキュリティ対策を使用[](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy)する必要があります。同じオリジン ポリシーと単純[な CORS](https://www.w3.org/TR/cors/)が必要です。

単純な CORS 実装では Cookie を使用できません。単純なメソッド (GET、HEAD、POST) のみをサポートします。 単純な CORS はフィールド名`Accept`、 `Accept-Language`、`Content-Language`の簡単なヘッダーを受け入れます。 コンテンツ タイプが 、 である場合は、単純な CORS でヘッダー `Content-Type` `application/x-www-form-urlencoded` `text/plain` を使用できます `multipart/form-data` 。

## <a name="storing-and-accessing-data"></a>データの格納およびアクセス

UI レスのカスタム関数内では、オブジェクトを使用してデータを格納およびアクセス `OfficeRuntime.storage` できます。 `Storage` は、UI レスのカスタム関数では使用できない [localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage)の代替手段を提供する、暗号化されていない永続的なキー値ストレージ システムです。 `Storage` ドメインごとに 10 MB のデータを提供します。 ドメインは、複数のアドインで共有できます。

`Storage` は共有ストレージ ソリューションとして機能することを意図しています。つまり、アドインの複数の部分が同じデータにアクセスできるようになります。 たとえば、ユーザー認証のトークンは、UI レスのカスタム関数と作業ウィンドウなどのアドイン UI 要素の両方からアクセスできるので、格納 `storage` できます。 同様に、2 つのアドインが同じドメイン (たとえば、 ) を共有している場合、情報の前後 `www.contoso.com/addin1` `www.contoso.com/addin2` の共有も許可されます `storage` 。 異なるサブドメインを持つアドインは、(たとえば、 ) の異なるインスタンスを持つ点に `storage` `subdomain.contoso.com/addin1` 注意 `differentsubdomain.contoso.com/addin2` してください。

`storage` は共有の場所として機能することから、キー値の組み合わせが書き換えられる可能性があることにご注意ください。

オブジェクトでは、次のメソッドを使用 `storage` できます。

- `getItem`
- `getItems`
- `setItem`
- `setItems`
- `removeItem`
- `removeItems`
- `getKeys`

> [!NOTE]
> すべての情報 (など) をクリアする方法はありません `clear` 。 代わりに、一度に複数のエントリを削除できる `removeItems` を使用してください。

### <a name="officeruntimestorage-example"></a>OfficeRuntime.storage の例

次のコード サンプルでは、キー `OfficeRuntime.storage.setItem` と値をに設定する関数を呼び出します `storage` 。

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a>その他の考慮事項

アドインで UI レスのカスタム関数のみを使用する場合は、UI レスのカスタム関数を使用してドキュメント オブジェクト モデル (DOM) にアクセスしたり、DOM に依存する jQuery のようなライブラリを使用したりすることはできません。

## <a name="next-steps"></a>次の手順
UI レスの [カスタム関数をデバッグする方法について説明します](custom-functions-debugging.md)。

## <a name="see-also"></a>関連項目

* [UI レスのカスタム関数を認証する](custom-functions-authentication.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
