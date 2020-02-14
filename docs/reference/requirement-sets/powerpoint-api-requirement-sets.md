---
title: PowerPoint JavaScript API の要件セット
description: ''
ms.date: 07/26/2019
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: 5bba2354cabba3c3ccd4ddf38d3e03c25a32b8a9
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950958"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a>PowerPoint JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

次の表は、PowerPoint の要件セット、それらの要件セットをサポートする Office ホスト アプリケーション、ビルド バージョンまたは一般提供開始日の一覧です。

|  要件セット  |  Windows での Office<br>(Office 365 サブスクリプションに接続)  |  Office on iPad<br>(Office 365 サブスクリプションに接続)  |  Office on Mac<br>(Office 365 サブスクリプションに接続)  | Office on the web |
|:-----|-----|:-----|:-----|:-----|:-----|
| PowerPointApi 1.1 | バージョン 1810 (ビルド 11001.20074) 以降 | 2.17 以降 | 16.19 以降 | 2018 年 10 月 |

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

Office のバージョンとビルド番号の詳細については、次を参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## <a name="powerpoint-javascript-api-11"></a>PowerPoint JavaScript API 1.1

PowerPoint JavaScript API 1.1 には、新しいプレゼンテーションを作成するための 1 つの API が含まれます。 API の詳細については、「[PowerPoint の JavaScript API](../../powerpoint/powerpoint-add-ins.md)」を参照してください。

## <a name="runtime-requirement-support-check"></a>ランタイム要件のサポートのチェック

実行時に、アドインは次を行うことによって、特定のホストが API 要件をサポートしているかどうかをチェックできます。

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
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
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

PowerPoint のほとんどのアドイン機能は、共通の API セットから取得されます。 共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [PowerPoint JavaScript API リファレンス ドキュメント](/javascript/api/powerpoint)
- [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)
