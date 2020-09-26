---
ms.date: 09/25/2020
description: 作業ウィンドウおよび特定の JavaScript ランタイムを使用しない Excel カスタム関数について説明します。
title: UI レス Excel カスタム関数のランタイム
localization_priority: Normal
ms.openlocfilehash: 94254dfb5a0d03b7c9fec392b2377aff91b58af4
ms.sourcegitcommit: b47318a24a50443b0579e05e178b3bb5433c372f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/25/2020
ms.locfileid: "48279514"
---
# <a name="runtime-for-ui-less-excel-custom-functions"></a>UI レス Excel カスタム関数のランタイム

作業ウィンドウを使用しないカスタム関数 (UI レスカスタム関数) は、計算のパフォーマンスを最適化するように設計された JavaScript ランタイムを使用します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

この JavaScript ランタイムは、UI を使用しない `OfficeRuntime` カスタム関数と作業ウィンドウでデータを格納するために使用できる名前空間の api へのアクセスを提供します。

## <a name="requesting-external-data"></a>外部データの要求

UI を使用しないカスタム関数内では、サーバーと対話するために HTTP 要求を発行する標準の web API である [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) や、 [XmlHttpRequest (xhr)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest)などの API を使用して外部データを要求できます。

UI を使用しない関数では、XmlHttpRequests を作成するときに追加のセキュリティ対策を使用する必要があることに注意してください。 [元のポリシー](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) と単純な [CORS](https://www.w3.org/TR/cors/)が必要です。

単純な CORS 実装は cookie を使用できず、simple メソッド (GET、HEAD、POST) のみをサポートしています。 単純な CORS はフィールド名`Accept`、 `Accept-Language`、`Content-Language`の簡単なヘッダーを受け入れます。 `Content-Type`コンテンツタイプが、、またはの場合は、単純な CORS のヘッダーを使用することもでき `application/x-www-form-urlencoded` `text/plain` `multipart/form-data` ます。

## <a name="storing-and-accessing-data"></a>データの格納およびアクセス

UI を使用しないカスタム関数では、オブジェクトを使用してデータを格納したり、データにアクセスしたりでき `OfficeRuntime.storage` ます。 `Storage` は、暗号化されていない、暗号化されていないキー値を持つ、永続的なストレージシステムです。これ [は、UI には](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage)ないカスタム関数では使用できません。 `Storage` ドメインごとに 10 MB のデータを提供します。 ドメインは複数のアドインで共有できます。

`Storage` は共有ストレージ ソリューションとして機能することを意図しています。つまり、アドインの複数の部分が同じデータにアクセスできるようになります。 たとえば、ユーザー認証のトークンは、 `storage` UI なしのカスタム関数と、作業ウィンドウなどのアドインの ui 要素の両方からアクセスできるため、に格納されます。 同様に、2つのアドインが同じドメイン (たとえば、など) を共有している場合は、 `www.contoso.com/addin1` `www.contoso.com/addin2` 情報を相互間で共有することもでき `storage` ます。 サブドメインが異なるアドインは、のインスタンスが異なることに注意 `storage` してください (例: `subdomain.contoso.com/addin1` `differentsubdomain.contoso.com/addin2` )。

`storage` は共有の場所として機能することから、キー値の組み合わせが書き換えられる可能性があることにご注意ください。

`storage` オブジェクトでは、以下の方法が利用可能です。

 - `getItem`
 - `getItems`
 - `setItem`
 - `setItems`
 - `removeItem`
 - `removeItems`
 - `getKeys`

> [!NOTE]
> すべての情報 (など) を消去する方法はありません `clear` 。 代わりに、一度に複数のエントリを削除できる `removeItems` を使用してください。

### <a name="officeruntimestorage-example"></a>一例

次のコードサンプルでは、関数を呼び出して `OfficeRuntime.storage.setItem` キーと値をに設定 `storage` します。

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

アドインで UI を使用しないカスタム関数のみが使用されている場合は、UI を使用しないカスタム関数を使用してドキュメントオブジェクトモデル (DOM) にアクセスしたり、DOM に依存している jQuery などのライブラリを使用したりすることができないことに注意してください。

## <a name="next-steps"></a>次の手順
[UI のないカスタム関数をデバッグ](custom-functions-debugging.md)する方法について説明します。

## <a name="see-also"></a>関連項目

* [UI レスのカスタム関数を認証する](custom-functions-authentication.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
