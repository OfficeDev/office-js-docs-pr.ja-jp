---
title: 画像強制型変換要件セット
description: Excel、PowerPoint、Word で Office アドインを使用して、画像の強制型変換の要件セットをサポートします。
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 59f6891182f47bed1b7e3b6aa69a30e941bce7cb
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094355"
---
# <a name="image-coercion-requirement-sets"></a>画像強制型変換要件セット

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 `Office.CoercionType.Image` は、メソッドを使用してデータを書き込むときに、image () への変換を有効に [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) します。 次のホストがサポートされています。

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

ImageCoercion 1.2 `Office.CoercionType.XmlSvg` は、メソッドを使用してデータを書き込むときに SVG 形式 () への変換を有効にし [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) ます。 次のホストがサポートされています。

- Windows 上の Excel (Microsoft 365 サブスクリプションに接続)
- Mac 上の Excel (Microsoft 365 サブスクリプションに接続)
- Windows 上の PowerPoint (Microsoft 365 サブスクリプションに接続されています)
- PowerPoint on Mac (Microsoft 365 サブスクリプションに接続)
- PowerPoint on the web
- Windows 上の Word (Microsoft 365 サブスクリプションに接続)
- Mac 上の Word (Microsoft 365 サブスクリプションに接続)
- Word on the web

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office のホストと API の要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
