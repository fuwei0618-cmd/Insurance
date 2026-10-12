/* 工作手冊設定
 * CLIENT_ID：Azure「Work」應用程式的「應用程式 (用戶端) 識別碼」。
 *   這不是密碼，放在公開的地方沒有關係；真正的資料只存在你的 OneDrive，而且是加密的。
 * FOLDER：資料放在 OneDrive › 應用程式 › Work › 這個資料夾裡。
 * AUTO_LOCK_MINUTES：多久沒操作就自動鎖定（要重新輸入密碼）。
 */
window.CRM_CONFIG = {
  CLIENT_ID: '',
  FOLDER: 'Cathay',
  AUTO_LOCK_MINUTES: 30,
};
