---
ms.date: 06/15/2022
description: 共有ランタイムExcel使用しないカスタム関数とその特定の JavaScript ランタイムについて説明します。
title: カスタム関数の JavaScript 専用ランタイム
ms.localizationpriority: medium
ms.openlocfilehash: 614e96937c769307b58e66943caa499f1f12d92c
ms.sourcegitcommit: d8fbe472b35c758753e5d2e4b905a5973e4f7b52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2022
ms.locfileid: "66229667"
---
# <a name="javascript-only-runtime-for-custom-functions"></a>カスタム関数の JavaScript 専用ランタイム

共有ランタイムを使用しないカスタム関数では、計算のパフォーマンスを最適化するように設計された JavaScript 専用ランタイムが使用されます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

この JavaScript ランタイムは、カスタム関数と作業ウィンドウ (別のランタイムで `OfficeRuntime` 実行される) でデータを格納するために使用できる名前空間内の API へのアクセスを提供します。

## <a name="request-external-data"></a>外部データを要求する

カスタム関数内では、[Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) などの API や、サーバーとやり取りする HTTP 要求を発行する標準 Web API である [XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) を使用して、外部データを要求できます。

XmlHttpRequests を作成する場合は、カスタム関数で追加のセキュリティ対策を使用する必要があることに注意してください。 [同じ配信元ポリシー](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) と単純な [CORS](https://www.w3.org/TR/cors/) が必要です。

単純な CORS 実装では Cookie を使用できず、単純なメソッド (GET、HEAD、POST) のみをサポートします。 単純な CORS はフィールド名`Accept`、 `Accept-Language`、`Content-Language`の簡単なヘッダーを受け入れます。 コンテンツ タイプ`application/x-www-form-urlencoded`が `Content-Type` 、`text/plain`、または `multipart/form-data`.

## <a name="store-and-access-data"></a>データのMicrosoft Storeとアクセス

共有ランタイムを使用しないカスタム関数内では、 [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) オブジェクトを使用してデータを格納してアクセスできます。 オブジェクトは `Storage` 、JavaScript 専用ランタイムを使用するカスタム関数では使用できない [localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) の代替手段を提供する、暗号化されていない永続的なキー値ストレージ システムです。 このオブジェクトは `Storage` 、ドメインごとに 10 MB のデータを提供します。 ドメインは、複数のアドインで共有できます。

オブジェクトは `Storage` 共有ストレージ ソリューションであり、アドインの複数の部分が同じデータにアクセスできることを意味します。 たとえば、ユーザー認証用のトークンは、カスタム関数 (JavaScript 専用ランタイムを使用) と作業ウィンドウ (フル Web ビュー ランタイムを使用) の両方からアクセスできるため、オブジェクトに格納 `Storage` できます。 同様に、2 つのアドインが同じドメイン (たとえば、) `www.contoso.com/addin2`を共有する場合、`www.contoso.com/addin1`オブジェクトを介して`Storage`情報を前後に共有することも許可されます。 サブドメインが異なるアドインには、異なるインスタンス `Storage` (例: `subdomain.contoso.com/addin1`, `differentsubdomain.contoso.com/addin2`) があることに注意してください。

オブジェクトは `Storage` 共有の場所であるため、キーと値のペアをオーバーライドできることを認識することが重要です。

オブジェクトでは、次のメソッドを `Storage` 使用できます。

- `getItem`
- `getItems`
- `setItem`
- `setItems`
- `removeItem`
- `removeItems`
- `getKeys`

> [!NOTE]
> すべての情報 (など `clear`) をクリアする方法はありません。 代わりに、一度に複数のエントリを削除できる `removeItems` を使用してください。

### <a name="officeruntimestorage-example"></a>OfficeRuntime.storage の例

次のコード サンプルでは、キーと値`storage`を`OfficeRuntime.storage.setItem`設定する関数を呼び出します。

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="next-steps"></a>次のステップ

[カスタム関数をデバッグ](custom-functions-debugging.md)する方法について説明します。

## <a name="see-also"></a>関連項目

* [共有ランタイムのないカスタム関数の認証](custom-functions-authentication.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
