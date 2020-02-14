---
title: OneNote JavaScript API の要件セット
description: ''
ms.date: 07/17/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 00bf9f23c307a6094345b753d7cccf1c10be7c32
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950965"
---
# <a name="onenote-javascript-api-requirement-sets"></a>OneNote JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

次の表は、OneNote の要件セット、それらの要件セットをサポートする Office ホスト アプリケーション、ビルド バージョンまたは一般提供開始日の一覧です。

|  要件セット  |  Office on the web |
|:-----|:-----|
| [OneNoteApi 1.1](/javascript/api/onenote?view=onenote-js-1.1)  | 2016 年 9 月 |  

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="onenote-javascript-api-11"></a>OneNote JavaScript API 1.1

OneNote JavaScript API 1.1 は、API の最初のバージョンです。 API について詳しくは、「[OneNote の JavaScript API のプログラミングの概要](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)」をご覧ください。

## <a name="runtime-requirement-support-check"></a>ランタイム要件のサポートのチェック

実行時に、アドインは次を行うことによって、特定のホストが API 要件をサポートしているかどうかをチェックできます。

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>マニフェストに基づく要件のサポートのチェック

アドインで必須の、重要な要件セットまたは API メンバーを指定するには、アドインのマニフェストで `Requirements` 要素を使用します。 Office ホストまたはプラットフォームが、`Requirements` 要素で指定した要件セットまたは API メンバーをサポートしない場合、アドインはそのホストまたはプラットフォームでは実行されず、[個人用アドイン] にも表示されません。

OneNoteApi 要件セット、バージョン 1.1 をサポートするすべての Office ホスト アプリケーションで読み込まれるアドインのコード例を以下に示します。

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
- [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)
