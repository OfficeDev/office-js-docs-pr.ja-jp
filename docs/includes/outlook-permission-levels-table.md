|**アクセス許可レベル</br>の正規名**|**XML マニフェスト名**|**Teams マニフェスト名**|**概要の説明**|
|:-----|:-----|:-----|:-----|
|**制限**|Restricted|MailboxItem.Restricted.User|エンティティを使用できますが、正規表現は使用できません。 |
|**アイテムの読み取り**|ReadItem|MailboxItem.Read.User|**制限** 付きで許可される内容に加えて、次のことが可能です。<ul><li>正規表現</li><li>Outlook アドイン API の読み取りアクセス</li><li>アイテムのプロパティとコールバック トークンの取得</li></ul> |
|**読み取り/書き込み項目**|ReadWriteItem|MailboxItem.ReadWrite.User|**読み取り項目** で許可される内容に加えて、次のことが可能です。<ul><li>`makeEwsRequestAsync` を除いた、完全な Outlook アドイン API のアクセス</li><li>アイテムのプロパティの設定</li></ul> |
|**メールボックスの読み取り/書き込み**|ReadWriteMailbox|Mailbox.ReadWrite.User|**読み取り/書き込み項目** で許可される内容に加えて、次のことが可能です。<ul><li>アイテムやフォルダーの作成、読み取り、書き込み</li><li>アイテムの送信</li><li>[makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) の呼び出し</li></ul> |

アクセス許可はマニフェストで宣言されます。 マークアップは、マニフェストの種類によって異なります。

- **XML マニフェスト**: 要素を使用します **\<Permissions\>** 。
- **Teams マニフェスト (プレビュー)**: "authorization.permissions.resourceSpecific" 配列内のオブジェクトの "name" プロパティを使用します。

> [!NOTE]
>
> - 追加の送信機能を使用するアドインには、補足的なアクセス許可が必要です。 XML マニフェストでは、 [ExtendedPermissions](/javascript/api/manifest/extendedpermissions) 要素でアクセス許可を指定します。 詳細については、「 [Outlook アドインに追加の送信を実装](../outlook/append-on-send.md)する」を参照してください。 Teams マニフェスト (プレビュー) では、"authorization.permissions.resourceSpecific" 配列の追加オブジェクトに **Mailbox.AppendOnSend.User** という名前でこのアクセス許可を指定します。
> - 共有フォルダーを使用するアドインには、追加のアクセス許可が必要です。 XML マニフェストでは、 [SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) 要素 `true`を . 詳細については、「 [Outlook アドインで共有フォルダーと共有メールボックスのシナリオを有効にする」を](../outlook/delegate-access.md)参照してください。 Teams マニフェスト (プレビュー) では、"authorization.permissions.resourceSpecific" 配列の追加オブジェクトに **Mailbox.SharedFolder** という名前でこのアクセス許可を指定します。
