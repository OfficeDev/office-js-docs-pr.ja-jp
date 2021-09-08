---
title: ランタイム ログを使用してアドインをデバッグする
description: ランタイム ログを使用してアドインをデバッグする方法を説明します。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: d191b2d7ac6135600bd6875ef7fbbced55caec8b
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937097"
---
# <a name="debug-your-add-in-with-runtime-logging"></a>ランタイム ログを使用してアドインをデバッグする

ランタイム ログを使用して、アドインのマニフェストやいくつかのインストール エラーをデバッグできます。 この機能は、リソース ID の不一致のような XSD スキーマ検証では検出されないマニフェストの問題を識別して修正するのに役立ちます。 ランタイム ログは、アドイン コマンドと Excel カスタム関数を実装するアドインのデバッグに特に有効です。

> [!NOTE]
> ランタイム ログ機能は、2016 Office以降のデスクトップで使用できます。

> [!IMPORTANT]
> ランタイムのログはパフォーマンスに影響します。アドイン マニフェストに関する問題をデバッグする必要がある場合にのみ有効にしてください。

## <a name="use-runtime-logging-from-the-command-line"></a>コマンド ラインからランタイム ログを使用する

コマンド ラインからランタイム ログを有効にするのが、このログ ツールを使用する最も簡単な方法です。 これは、npm@5.2.0+ の一部として既定で提供される npx を使用します。 以前のバージョンの [npm](https://www.npmjs.com/) を使用している場合は、[Windows でのランタイム ログ](#runtime-logging-on-windows)の手順か [Mac でのランタイム ログ](#runtime-logging-on-mac)の手順、または [npx のインストール](https://www.npmjs.com/package/npx)をお試しください。

- ランタイムのログを有効にするには、以下を実行します。

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable
    ```

- 特定のファイルに対してのみランタイム ログを有効にするには、ファイル名と同じコマンドを使用します。

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable [filename.txt]
    ```

- ランタイム ログを無効にするには、以下を実行します。

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --disable
    ```

- ランタイム ログが有効になっているかどうかを表示するには、以下を実行します。

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log
    ```

- ランタイム ログのコマンド ライン内にヘルプを表示するには、以下を実行します。

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --help
    ```

## <a name="runtime-logging-on-windows"></a>Windows でのランタイム ログ

1. Office 2016 デスクトップのビルド **16.0.7019** 以降を実行していることを確認します。

2. `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\` の下に `RuntimeLogging` レジストリ キーを追加します。

    [!include[Developer registry key](../includes/developer-registry-key.md)]

3. **RuntimeLogging** キーの既定値にログを書き込むファイルの完全なパスを設定します。 例については、[EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip) を参照してください。

    > [!NOTE]
    > ログ ファイルが書き込まれるディレクトリが既に存在しており、書き込みアクセス許可がある必要があります。

レジストリは次の図のようになります。 この機能を無効にするには、`RuntimeLogging` キーをレジストリから削除します。

![RuntimeLogging レジストリ キーを使用したレジストリ エディターのスクリーンショット。](../images/runtime-logging-registry.png)

## <a name="runtime-logging-on-mac"></a>Mac でのランタイム ログ

1. Office 2016 デスクトップのビルド **16.27** (19071500) 以降を実行していることを確認します。

2. **ターミナル** を開き、`defaults`コマンドを使用してランタイム ログの優先度を設定します。

    ```command&nbsp;line
    defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
    ```

    `<bundle id>`は、ランタイム ログを有効にするホストを識別します。 `<file_name>`は、ログが書き込まれるテキスト ファイルの名前です。

    対応 `<bundle id>` するアプリケーションのランタイム ログを有効にするには、次のいずれかの値に設定します。

    - `com.microsoft.Word`
    - `com.microsoft.Excel`
    - `com.microsoft.Powerpoint`
    - `com.microsoft.Outlook`

次の例では、Word のランタイム ログを有効にし、ログ ファイルを開きます。

```command&nbsp;line
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string "runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```

> [!NOTE]
> ランタイム ログを有効にするには、`defaults`コマンドを実行した後に Office を再起動する必要があります。

ランタイム ログを無効にするには、`defaults delete`コマンドを使用します。

```command&nbsp;line
defaults delete <bundle id> CEFRuntimeLoggingFile
```

次の例では、Word のランタイム ログを無効にします。

```command&nbsp;line
defaults delete com.microsoft.Word CEFRuntimeLoggingFile
```

## <a name="use-runtime-logging-to-troubleshoot-issues-with-your-manifest"></a>ランタイム ログを使用してマニフェストでの問題のトラブルシューティングを行う

ランタイムのログを使用してアドインの読み込みに関する問題のトラブルシューティングを行うには、次のようにします。

1. テスト用に[アドインをサイドロード](sideload-office-add-ins-for-testing.md)します。

    > [!NOTE]
    > ログ ファイルのメッセージ数を最小限に抑えるため、テストするアドインのみをサイドロードすることをお勧めします。

2. 何も起こらず、アドインが表示されない (アドイン ダイアログ ボックスにも表示されない) 場合は、ログ ファイルを開きます。

3. ログ ファイルでアドインの ID を検索します。ID はマニフェストで定義します。ログ ファイルでは、この ID には `SolutionId` というラベルが付いています。

## <a name="known-issues-with-runtime-logging"></a>ランタイムのログに関する既知の問題

混乱を招くメッセージまたは正しく分類されていないメッセージがログ ファイルに書き込まれることがあります。たとえば次のような場合です。

- メッセージ "`Medium Current host not in add-in's host list`" に続く "`Unexpected Parsed manifest targeting different host`" は、誤ってエラーとして分類されています。

- SolutionId が含まれていないメッセージ "`Unexpected Add-in is missing required manifest fields    DisplayName`" は、多くの場合、エラーはデバッグ対象のアドインと関係ありません。

- `Monitorable` メッセージは、システムの観点からのエラーと予想されます。場合によっては、スキップされたがマニフェスト失敗の原因にはならなかったスペル ミスのある要素のような、マニフェストの問題を示していることがあります。

## <a name="see-also"></a>関連項目

- [Office アドインの XML マニフェスト](../develop/add-in-manifests.md)
- [Office アドインのマニフェストを検証する](troubleshoot-manifest.md)
- [Office のキャッシュをクリアする](clear-cache.md)
- [テスト用に Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)
- [Office アドインをデバッグする](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
