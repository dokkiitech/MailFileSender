function saveReceiptsToDrive() {
  const query = 'from:invoice+statements+acct_1C3f9UEkfyRW2REo@stripe.com filename:Receipt';
  const threads = GmailApp.search(query);
  const folder = DriveApp.getFolderById('FolderID'); // 保存するフォルダのIDを指定してください
  const savedFileNames = new Set();

  // 既存ファイルの名前を保存
  const savedFiles = folder.getFiles();
  while (savedFiles.hasNext()) {
    const file = savedFiles.next();
    savedFileNames.add(file.getName());
  }

  const newFiles = [];

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      const attachments = message.getAttachments();
      attachments.forEach(attachment => {
        const fileName = attachment.getName();
        if (fileName.startsWith('Receipt') && !savedFileNames.has(fileName)) {
          folder.createFile(attachment);
          savedFileNames.add(fileName);
          newFiles.push(fileName);
        }
      });
    });
  });

  // 保存完了通知を送信
  if (newFiles.length > 0) {
    const recipients = ['example@example.com']; // 通知を送信するメールアドレスを指定してください アドレスを追加するときはカンマ区切りでアドレスを追加 ['example@example.com' ,'example@example.com']
    const subject = 'Cirport領収書を保存しました。';
    const body = 'Cirportの領収書を保存しました。\n\n' +
                 '以下のリンクよりご確認の上、指定口座にお振り込みください。\n' +
                 'googleDrive_FolderURL\n\n' + //フォルダーのURLを指定してください
                 '追加されたファイル:\n' + newFiles.join('\n');
    recipients.forEach(recipient => MailApp.sendEmail(recipient, subject, body));
  }
}
