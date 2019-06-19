---
title: Office アドインを既存の COM アドインと互換できるようにする
description: Office アドインと同等の COM アドインの互換性を有効にする
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: a18adb9841a9580d77c5110a0346f365e38e3746
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059721"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in-preview"></a>既存の COM アドインと互換性のある Office アドインを作成する (プレビュー)

既存の COM アドインがある場合は、Office アドインで同等の機能を構築できます。これにより、web や Office on the Mac 上の他のプラットフォーム上でソリューションを実行することが可能になります。 場合によっては、Office アドインが、対応する COM アドインで使用可能なすべての機能を提供できないことがあります。 このような状況では、対応する Office アドインが提供するよりも、COM アドインによって Windows のユーザーの利便性が向上することがあります。

同等の COM アドインがユーザーのコンピューターに既にインストールされている場合に office アドインを構成すると、office アドインではなく、Windows が COM アドインを実行するようになります。 COM アドインは、Office がユーザーのコンピューターにインストールされているものに応じて、COM アドインと Office アドインをシームレスに移行するため、"同等" と呼ばれます。

> [!NOTE]
> この機能は現在プレビュー段階で、運用環境での使用はサポートされていません。 これは、Excel、Word、および PowerPoint のバージョン16.0.11629.20214 以降で使用できます。 このビルドにアクセスするには、Office 365 サブスクリプションを用意し、 **insider**レベルで[office insider](https://products.office.com/office-insider)プログラムに参加する必要があります。

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a>マニフェストで同等の COM アドインを指定する

Office アドインと COM アドインの互換性を有効にするには、Office アドインの[マニフェスト](add-in-manifests.md)で同等の COM アドインを特定します。 その後、office アドインの両方がインストールされている場合は、Windows で office アドインではなく COM アドインが使用されます。

次の例は、COM アドインを同等のアドインとして指定するマニフェストの一部を示しています。 `ProgId`要素の値は COM アドインを識別し、 `EquivalentAddins`要素は終了`VersionOverrides`タグの直前に配置する必要があります。

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  <EquivalentAddins>
</VersionOverrides>
```

> [!TIP]
> COM アドインと XLL UDF の互換性の詳細については、「 [xll ユーザー定義関数と互換性のあるカスタム関数を作成する](../excel/make-custom-functions-compatible-with-xll-udf.md)」を参照してください。

## <a name="equivalent-behavior-for-users"></a>ユーザーの同等の動作

Office アドインマニフェストで同等の COM アドインが指定されている場合、Windows 上の Office では、対応する COM アドインがインストールされている場合、Office アドインのユーザーインターフェイス (UI) は表示されません。 Office は、Office アドインのリボンボタンを非表示にし、インストールを妨げることはありません。 そのため、Office アドインは引き続き UI 内の次の場所に表示されます。

- [**個人用アドイン] の**下
- リボンマネージャーのエントリとして

> [!NOTE]
> マニフェストで同等の COM アドインを指定しても、web または Office for Mac の Office などの他のプラットフォームには影響しません。

次のシナリオでは、ユーザーが Office アドインを取得する方法によって実行される処理について説明します。

### <a name="appsource-acquisition-of-an-office-add-in"></a>Office アドインの AppSource 取得

ユーザーが AppSource から Office アドインを取得し、対応する COM アドインが既にインストールされている場合、Office は次のようになります。

1. Office アドインをインストールします。
2. リボンで Office アドイン UI を非表示にします。
3. [COM アドイン] リボンボタンをポイントするユーザーの呼び出しを表示します。

### <a name="centralized-deployment-of-office-add-in"></a>Office アドインの一元展開

管理者が一元展開を使用して Office アドインをテナントに展開しており、対応する COM アドインが既にインストールされている場合、ユーザーは Office を再起動して変更を表示する必要があります。 Office を再起動すると、次のようになります。

1. Office アドインをインストールします。
2. リボンで Office アドイン UI を非表示にします。
3. [COM アドイン] リボンボタンをポイントするユーザーの呼び出しを表示します。

### <a name="document-shared-with-embedded-office-add-in"></a>埋め込まれた Office アドインと共有されたドキュメント

ユーザーが COM アドインをインストールしていて、Office アドインが埋め込まれた共有ドキュメントを取得した場合、Office は次のようになります。

1. Office アドインを信頼するかどうかをユーザーに確認します。
2. 信頼できる場合は、Office アドインがインストールされます。
3. リボンで Office アドイン UI を非表示にします。

## <a name="other-com-add-in-behavior"></a>その他の COM アドインの動作

ユーザーが同等の COM アドインをアンインストールした場合は、Windows の Office によって Office アドインの UI が復元されます。

Office アドインに対応する COM アドインを指定した後、office アドインの更新プログラムの処理を停止します。 Office アドインの最新の更新プログラムを入手するには、まず COM アドインをアンインストールする必要があります。

## <a name="see-also"></a>関連項目

- [カスタム関数を XLL ユーザー定義関数と互換性を持つようにする](../excel/make-custom-functions-compatible-with-xll-udf.md)
