---
title: 画像強制型変換要件セット
description: 複数のアドインを使用した Image Coercion 要件セットOffice、Excel、Word PowerPointサポート。
ms.date: 02/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: d605a8ac5d36825ffdfa9648578f7d252340325e1021c7663667778f4427ef6a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57090649"
---
# <a name="image-coercion-requirement-sets"></a>画像強制型変換要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 では、メソッドを使用してデータを書き込むときにイメージ ( `Office.CoercionType.Image` ) への変換が有効 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) です。 次のアプリケーションがサポートされています。

- Excel 2013 以降のWindows
- Excel 2016以降の Mac
- Excel on iPad
- OneNote on the web
- PowerPoint 2013 以降のWindows
- PowerPoint 2016以降の Mac
- PowerPoint on the web
- PowerPoint on iPad
- Word on Windows (Word 2013 以降)
- Word on Mac (Word 2016 以降)
- Word on the web
- Word on iPad

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1.2 では、メソッドを使用してデータを書き込むときに SVG 形式 ( `Office.CoercionType.XmlSvg` ) に変換 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) できます。 次のアプリケーションがサポートされています。

- ExcelオンWindows (サブスクリプションに接続Microsoft 365)
- Excel Mac (サブスクリプションに接続Microsoft 365)
- PowerPoint (WindowsサブスクリプションにMicrosoft 365)
- PowerPoint (サブスクリプションにMicrosoft 365)
- PowerPoint on the web
- Word on Windows (サブスクリプションにMicrosoft 365)
- Mac 上の Word (サブスクリプションに接続Microsoft 365)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
