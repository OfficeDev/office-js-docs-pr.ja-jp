---
title: 画像強制型変換要件セット
description: Excel、PowerPoint、Word で Office アドインを使用して、画像の強制型変換の要件セットをサポートします。
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 83817bfc7cf8a193138a805b0e90b4357d605801
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596971"
---
# <a name="image-coercion-requirement-sets"></a>画像強制型変換要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 は、メソッドを使用し`Office.CoercionType.Image`てデータを書き込むときに[`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 、image () への変換を有効にします。 次のホストがサポートされています。

- Excel 2013 以降
- Excel 2016 以降 (Mac)
- Excel on iPad
- OneNote on the web
- PowerPoint 2013 以降
- PowerPoint 2016 以降の Mac
- PowerPoint on the web
- PowerPoint on iPad
- Word on Windows (Word 2013 以降)
- Word on Mac (Word 2016 以降)
- Word on the web
- Word on iPad

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1.2 は、メソッドを使用し`Office.CoercionType.XmlSvg`てデータを書き込むときに[`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) SVG 形式 () への変換を有効にします。 次のホストがサポートされています。

- Windows 上の Excel (Office 365 サブスクリプションに接続)
- Excel on Mac (Office 365 サブスクリプションに接続)
- Windows 上の PowerPoint (Office 365 サブスクリプションに接続)
- PowerPoint on Mac (Office 365 サブスクリプションに接続されている)
- PowerPoint on the web
- Windows 上の Word (Office 365 サブスクリプションに接続)
- Mac 上の Word (Office 365 サブスクリプションに接続されている)
- Word on the web

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office のホストと API の要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
