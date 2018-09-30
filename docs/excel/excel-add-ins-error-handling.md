---
title: エラー処理
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 23a70b1d66befb971c3c1394eb9162c19f2ee176
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348087"
---
# <a name="error-handling"></a>エラー処理

Excel JavaScript API を使用してアドインを作成する場合は、実行時エラーを考慮するために、エラー処理ロジックを含めます。 これは、API の非同期性のために重要になります。

> [!NOTE]
> **sync()** メソッドと Excel JavaScript API の非同期性の詳細については、「[Excel JavaScript API の中心概念](excel-add-ins-core-concepts.md)」を参照してください。

## <a name="best-practices"></a>ベスト プラクティス

このドキュメントのコード サンプルでは、`Excel.run` へのすべての呼び出しが、`Excel.run` 内で発生したエラーを検出するために `catch` 文に付随して起こっていることがわかります。 Excel JavaScript Api を使用してアドインを作成する場合は、同じパターンを使用することをお勧めします。

```js
Excel.run(function (context) { 
  
  // Excel JavaScript API calls here

  // Await the completion of context.sync() before continuing.
  return context.sync()
    .then(function () {
      console.log("Finished!");
    })
}).catch(errorHandlerFunction);     
```

## <a name="api-errors"></a>API エラー 

Excel JavaScript API 要求が正常に実行されない場合、API は次のプロパティを含むエラー オブジェクトを返します。 

- **code**: エラー メッセージの `code` プロパティには、`OfficeExtension.ErrorCodes` または `Excel.ErrorCodes` リストの一部である文字列が含まれます。 たとえば、エラー コード "InvalidReference" は、参照が指定された操作に対して有効でないことを示します。 エラー コードはローカライズされません。 

- **message**: エラー メッセージの `message` プロパティには、ローカライズされた文字列のエラーの概要が含まれます。 このエラー メッセージは、エンド ユーザーが使用するためのものではありません。エラー コードと適切なビジネス ロジックを使用して、アドインがエンド ユーザーに表示するエラー メッセージを判断する必要があります。

- **debugInfo**: 存在する場合、エラー メッセージの `debugInfo` プロパティは、エラーの根本原因を理解するために使用できる追加情報を提供します。 

> [!NOTE]
> `console.log()` を使用してエラー メッセージをコンソールに出力すると、それらのメッセージはサーバー上でのみ表示されます。 これらのエラー メッセージが、アドインの作業ウィンドウやホスト アプリケーション内のいずれかの場所で、エンド ユーザーに対して表示されることはありません。

## <a name="see-also"></a>関連項目

- [Excel JavaScript API の中心概念](excel-add-ins-core-concepts.md)
- [OfficeExtension.Error オブジェクト (Excel の JavaScript API)](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
