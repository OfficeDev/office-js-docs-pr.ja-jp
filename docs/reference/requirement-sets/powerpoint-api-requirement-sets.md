---
title: PowerPoint JavaScript API の要件セット
description: ''
ms.date: 07/26/2019
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 4f64654a4130cc0d4bf96d9c59e364e77c808748
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/31/2019
ms.locfileid: "35941149"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a>PowerPoint JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

次の表に、PowerPoint の要件セット、それらの要件セットをサポートする Office ホストアプリケーション、ビルドバージョンまたは使用可能な日付を示します。

|  要件セット  |  Windows での Office<br>(Office 365 サブスクリプションに接続)  |  Office on iPad<br>(Office 365 サブスクリプションに接続)  |  Mac 版 Office<br>(Office 365 サブスクリプションに接続)  | Web 上の Office |
|:-----|-----|:-----|:-----|:-----|:-----|
| PowerPointApi 1.1 | バージョン 1810 (ビルド 11001.20074) 以降 | 2.17 以降 | 16.19 以降 | 2018 年 10 月 |

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

Office のバージョンとビルド番号の詳細については、以下を参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## <a name="powerpoint-javascript-api-11"></a>PowerPoint JavaScript API 1.1

PowerPoint JavaScript API 1.1 には、新しいプレゼンテーションを作成するための単一の API が含まれています。 API の詳細については、「 [JAVASCRIPT api For PowerPoint](../../powerpoint/powerpoint-add-ins.md)」を参照してください。

## <a name="runtime-requirement-support-check"></a>ランタイム要件のサポートのチェック

実行時に、アドインは、次の手順に従って、特定のホストが API 要件セットをサポートしているかどうかを確認できます。

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>マニフェストに基づく要件のサポートのチェック

アドインマニフェスト`Requirements`の要素を使用して、アドインが使用する必要がある重要な要件セットまたは API メンバーを指定します。 Office ホストまたはプラットフォームが、 `Requirements`要素で指定されている要件セットや API メンバーをサポートしていない場合、アドインはそのホストまたはプラットフォームでは実行されず、アドインには表示されません。

OneNoteApi 要件セット、バージョン 1.1 をサポートするすべての Office ホスト アプリケーションで読み込まれるアドインのコード例を以下に示します。

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

PowerPoint アドインのほとんどの機能は、共通 API セットから取得されます。 共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [PowerPoint JavaScript API リファレンスドキュメント](/javascript/api/powerpoint)
- [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)
