---
title: Office アドインに既存の COM アドインとの互換性をもたせる
description: アドインと同等の COM アドインOffice互換性を有効にする。
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: b5235255987cc6a644491bc548b92701b350a179
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836855"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a>Office アドインに既存の COM アドインとの互換性をもたせる

既存の COM アドインがある場合は、Office アドインで同等の機能を構築して、web や Mac 上の Office などの他のプラットフォームでソリューションを実行できます。 場合によっては、Office COM アドインで使用可能なすべての機能を提供できない場合があります。 このような状況では、COM アドインは、アドインが提供できる対応する機能よりも、Windows でのユーザー エクスペリエンスOffice向上することがあります。

Office アドインを構成して、同等の COM アドインが既にユーザーのコンピューターにインストールされている場合、Windows 上の Office が Office アドインではなく COM アドインを実行します。 COM アドインは、Office がユーザーのコンピューターにインストールされている COM アドインと Office アドインの間でシームレスに切り替わるため、「同等」と呼ばれる。

> [!NOTE]
> この機能は、Microsoft 365 サブスクリプションに接続されている場合、次のプラットフォームでサポートされます。
>
> - Web 上の Excel、Word、および PowerPoint
> - Windows 上の Excel、Word、および PowerPoint (バージョン 1904 以降)
> - Mac 上の Excel、Word、および PowerPoint (バージョン 13.329 以降)
> - Outlook on Windows (バージョン 2102 以降)

## <a name="specify-an-equivalent-com-add-in"></a>同等の COM アドインを指定する

### <a name="manifest"></a>マニフェスト

> [!IMPORTANT]
> Excel、PowerPoint、および Word に適用されます。 Outlook のサポートは近日公開予定です。

Office アドインと COM アドイン間の互換性を有効にするには、Office アドインのマニフェストで同等の COM アドインを[](add-in-manifests.md)識別します。 次Office Windows では、両方がインストールされている場合は、Officeアドインの代わりに COM アドインが使用されます。

次の例は、COM アドインを同等のアドインとして指定するマニフェストの部分を示しています。 要素の値は COM アドインを識別し `ProgId` [、EquivalentAddins](../reference/manifest/equivalentaddins.md) 要素は終了タグの直前に配置する必要 `VersionOverrides` があります。

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

> [!TIP]
> COM アドインと XLL UDF の互換性については、「カスタム関数を XLL ユーザー定義関数と互換性のあるものにする [」を参照してください](../excel/make-custom-functions-compatible-with-xll-udf.md)。

### <a name="group-policy"></a>グループ ポリシー

> [!IMPORTANT]
> Outlook にのみ適用されます。

Outlook Web アドインと COM/VSTO アドインの互換性を宣言するには、グループ ポリシーで同等の COM アドインを識別し、ユーザーのコンピューターで構成することにより、同等の COM または **VSTO** アドインがインストールされている Outlook Web アドインを非アクティブ化します。 次に、Outlook on Windows では、両方がインストールされている場合は、Web アドインの代わりに COM アドインを使用します。

1. ツールの [インストール手順に](https://www.microsoft.com/download/details.aspx?id=49030)注意を払って、最新の管理用テンプレート ツール **をダウンロードします**。
1. ローカル グループ ポリシー エディター **(gpedit.msc) を開きます**。
1. [ユーザー **構成]**  >  **[管理用テンプレート**]   >  **[Microsoft Outlook 2016**  >  **その他] に移動します**。
1. 同等の **COM または VSTO** アドインがインストールされている Outlook Web アドインを非アクティブ化する設定を選択します。
1. リンクを開き、ポリシー設定を編集します。
1. ダイアログの **Outlook Web アドインで非アクティブ化するには、次の操作を行います**。
    1. [ **値の名前]** を `Id` Web アドインのマニフェストで見つかった名前に設定します。 **重要**: *中かっこ* をエントリの周囲 `{}` に追加しない。
    1. Value **を** 同等 `ProgId` の COM/VSTO アドインの値に設定します。
    1. **[OK] を** 選択して更新プログラムを有効にします。
    ![ダイアログ "非アクティブ化する Outlook Web アドイン" を示すスクリーンショット](../images/outlook-deactivate-gpo-dialog.png)

## <a name="equivalent-behavior-for-users"></a>ユーザーと同等の動作

同等の [COM](#specify-an-equivalent-com-add-in)アドインを指定すると、windows 上の Office は、同等の COM アドインがインストールされている場合、Office アドインのユーザー インターフェイス (UI) は表示されません。 Officeアドインのリボン ボタンのみを非表示Officeし、インストールを妨げる必要があります。 したがって、Officeアドインは UI 内の次の場所に表示されます。

- [ **自分のアドイン] の下**
- リボン マネージャーのエントリとして (Excel、Word、および PowerPoint のみ)

> [!NOTE]
> マニフェストで同等の COM アドインを指定すると、web 上や Mac 上の Officeなどの他のプラットフォームには影響しません。

次のシナリオでは、ユーザーがアドインを取得する方法に応Office説明します。

### <a name="appsource-acquisition-of-an-office-add-in"></a>AppSource によるアドインOffice取得

ユーザーが AppSource から Officeアドインを取得し、同等の COM アドインが既にインストールされている場合は、次Officeします。

1. アドインOfficeインストールします。
2. リボンでOfficeアドイン UI を非表示にします。
3. COM アドイン リボン ボタンをポイントするユーザーの呼び出しを表示します。

### <a name="centralized-deployment-of-office-add-in"></a>アドインのOffice展開

管理者が集中展開を使用して Office アドインをテナントに展開し、同等の COM アドインが既にインストールされている場合、ユーザーは変更を表示する前に Office を再起動する必要があります。 再起動Office、次のコマンドが実行されます。

1. アドインOfficeインストールします。
2. リボンでOfficeアドイン UI を非表示にします。
3. COM アドイン リボン ボタンをポイントするユーザーの呼び出しを表示します。

### <a name="document-shared-with-embedded-office-add-in"></a>埋め込みアドインと共有Officeドキュメント

ユーザーが COM アドインをインストールし、埋め込み Office アドインを含む共有ドキュメントを取得した場合、そのユーザーがドキュメントを開いた場合、次のOfficeされます。

1. ユーザーにアドインを信頼Office求めるメッセージを表示します。
2. 信頼できる場合は、Officeアドインがインストールされます。
3. リボンでOfficeアドイン UI を非表示にします。

## <a name="other-com-add-in-behavior"></a>その他の COM アドインの動作

### <a name="excel-powerpoint-word"></a>Excel、PowerPoint、Word

ユーザーが同等の COM アドインをアンインストールした場合は、Windows Officeアドイン UI Office復元します。

カスタム アドインに同等の COM アドインを指定したOffice、Officeの更新プログラムの処理Office停止します。 アドインの最新の更新プログラムOffice、ユーザーはまず COM アドインをアンインストールする必要があります。

### <a name="outlook"></a>Outlook

対応する Web アドインを無効にするには、Outlook の起動時に COM/VSTO アドインを接続する必要があります。

その後の Outlook セッション中に COM/VSTO アドインが切断された場合、Outlook が再起動されるまで、Web アドインは無効なままである可能性があります。

## <a name="see-also"></a>関連項目

- [カスタム関数を XLL ユーザー定義関数と互換性のあるものにする](../excel/make-custom-functions-compatible-with-xll-udf.md)
