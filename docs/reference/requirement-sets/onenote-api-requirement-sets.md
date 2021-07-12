---
title: OneNote JavaScript API の要件セット
description: OneNote JavaScript API の要件セットについて説明します。
ms.date: 08/24/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: ecdb26edca54758540688ba03b1d9c1eec14e739
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350191"
---
# <a name="onenote-javascript-api-requirement-sets"></a>OneNote JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

次の表は、OneNote の要件セット、それらの要件セットをサポートする Office クライアント アプリケーション、ビルド バージョンまたは一般提供開始日の一覧です。

|  要件セット  |  Office on the web |
|:-----|:-----|
| [OneNoteApi 1.1](/javascript/api/onenote?view=onenote-js-1.1&preserve-view=true)  | 2016 年 9 月 |  

## <a name="onenote-javascript-api-11"></a>OneNote JavaScript API 1.1

OneNote JavaScript API 1.1 は、API の最初のバージョンです。API について詳しくは、「[OneNote の JavaScript API のプログラミングの概要](../../onenote/onenote-add-ins-programming-overview.md)」をご覧ください。

## <a name="runtime-requirement-support-check"></a>ランタイム要件のサポートのチェック

実行時に、アドインは次を行うことによって、特定の Office アプリケーションが API 要件をサポートしているかどうかをチェックできます。

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>マニフェストに基づく要件のサポートのチェック

アドインで必須の、重要な要件セットまたは API メンバーを指定するには、アドインのマニフェストで `Requirements` 要素を使用します。Office アプリケーションまたはプラットフォームが、`Requirements` 要素で指定した要件セットまたは API メンバーをサポートしない場合、アドインはそのアプリケーションまたはプラットフォームでは実行されず、[個人用アドイン] にも表示されません。

OneNoteApi 要件セット、バージョン 1.1 をサポートするすべての Office クライアント アプリケーションで読み込まれるアドインのコード例を以下に示します。

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [OneNote JavaScript API リファレンス ドキュメント](/javascript/api/onenote)
- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
