---
title: 画像強制型変換要件セット
description: Excel、PowerPoint、および Word Officeアドインを使用した Image Coercion 要件セットのサポート。
ms.date: 02/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 52ce46a46580500f5a292bf898674d4798378319
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505529"
---
# <a name="image-coercion-requirement-sets"></a>画像強制型変換要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 では、メソッドを使用してデータを書き込むときにイメージ ( `Office.CoercionType.Image` ) への変換が有効 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) です。 次のアプリケーションがサポートされています。

- Windows 上の Excel 2013 以降
- Mac での Excel 2016 以降
- Excel on iPad
- OneNote on the web
- Windows の PowerPoint 2013 以降
- Mac の PowerPoint 2016 以降
- PowerPoint on the web
- PowerPoint on iPad
- Word on Windows (Word 2013 以降)
- Word on Mac (Word 2016 以降)
- Word on the web
- Word on iPad

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1.2 では、メソッドを使用してデータを書き込むときに SVG 形式 ( `Office.CoercionType.XmlSvg` ) に変換 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) できます。 次のアプリケーションがサポートされています。

- Excel on Windows (Microsoft 365 サブスクリプションに接続)
- Excel on Mac (Microsoft 365 サブスクリプションに接続)
- PowerPoint on Windows (Microsoft 365 サブスクリプションに接続)
- PowerPoint on Mac (Microsoft 365 サブスクリプションに接続)
- PowerPoint on the web
- Word on Windows (Microsoft 365 サブスクリプションに接続)
- Mac 上の Word (Microsoft 365 サブスクリプションに接続)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
