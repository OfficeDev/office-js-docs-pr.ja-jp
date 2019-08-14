---
title: 画像強制の要件セット
description: Excel、PowerPoint、Word で Office アドインを使用して、画像の強制型変換の要件セットをサポートします。
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 9d622c827315f6657cf0fddaace33968bd634d64
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395674"
---
# <a name="image-coercion-requirement-sets"></a>画像強制の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 は、メソッドを使用し`Office.CoercionType.Image`てデータを書き込むときに[`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 、image () への変換を有効にします。 次のホストがサポートされています。

- Excel 2013 以降
- Excel 2016 以降 (Mac)
- IPad の Excel
- Web 上の OneNote
- PowerPoint 2013 以降
- PowerPoint 2016 以降の Mac
- PowerPoint on the web
- IPad の PowerPoint
- Word 2013 以降 (Windows)
- Word 2016 以降の Mac
- Web 上の Word
- iPad の Word

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1.2 は、メソッドを使用し`Office.CoercionType.XmlSvg`てデータを書き込むときに[`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) SVG 形式 () への変換を有効にします。 次のホストがサポートされています。

- Windows 上の Excel (Office 365 サブスクリプションに接続)
- Excel on Mac (Office 365 サブスクリプションに接続)
- Windows 上の PowerPoint (Office 365 サブスクリプションに接続)
- PowerPoint on Mac (Office 365 サブスクリプションに接続されている)
- PowerPoint on the web
- Windows 上の Word (Office 365 サブスクリプションに接続)
- Mac 上の Word (Office 365 サブスクリプションに接続されている)
- Web 上の Word

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)
