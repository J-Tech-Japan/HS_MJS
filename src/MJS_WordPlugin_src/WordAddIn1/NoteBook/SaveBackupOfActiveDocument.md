# SaveBackupOfActiveDocument

## �T�v
ThisAddIn.cs �� `SaveBackupOfActiveDocument` ���\�b�h�́A�A�N�e�B�u��Word�����̃o�b�N�A�b�v�t�@�C���i_backup�t���j�������I�ɍ쐬���鏈�����s���܂��B

## �ڍׂȏ����菇
1. �A�N�e�B�u�ȕ����iActiveDocument�j�����݂��A���t�@�C���p�X����łȂ��ꍇ�̂ݏ����𑱍s���܂��B
2. �t�@�C�����̖������u_backup�v�ł���Ή��������I�����܂��i�o�b�N�A�b�v�t�@�C�����g�̓o�b�N�A�b�v���Ȃ��j�B
3. �o�b�N�A�b�v�t�@�C�����i���̃t�@�C�����{_backup�j�𐶐����܂��B
4. �o�b�N�A�b�v�t�@�C�������ɑ��݂���ꍇ�͉��������I�����܂��B
5. �o�b�N�A�b�v�t�@�C�������݂��Ȃ��ꍇ�́A���̃t�@�C�����R�s�[���ăo�b�N�A�b�v�t�@�C�����쐬���܂��B

## �ړI
- �A�N�e�B�u�ȕ����̃o�b�N�A�b�v�������I�ɍ쐬���A���t�@�C���̕ی�╜����e�Ղɂ��܂��B
- �ҏW�O�̏�Ԃ�ێ��������ꍇ��A�둀��ɂ��f�[�^������h�����߂ɗL���ł��B

## ���ӓ_
- ���Ƀo�b�N�A�b�v�t�@�C�������݂���ꍇ�͐V���ɍ쐬���܂���B
- �o�b�N�A�b�v�Ώۂ́u_backup�v�ŏI���Ȃ��t�@�C���݂̂ł��B
- �t�@�C�����쎞�ɗ�O�����������ꍇ�̓G���[���b�Z�[�W���\������܂��B

```CSharp
// ThisAddIn.cs
private void SaveBackupOfActiveDocument()
{
    try
    {
        var doc = this.Application.ActiveDocument;
        if (doc != null && !string.IsNullOrEmpty(doc.FullName))
        {
            string originalPath = doc.FullName;
            string dir = Path.GetDirectoryName(originalPath);
            string name = Path.GetFileNameWithoutExtension(originalPath);
            string ext = Path.GetExtension(originalPath);

            // ������ "_backup" �̏ꍇ�͉������Ȃ�
            if (name.EndsWith("_backup", StringComparison.OrdinalIgnoreCase))
                return;

            string backupName = name + "_backup" + ext;
            string backupPath = Path.Combine(dir, backupName);

            // ���łɃo�b�N�A�b�v�t�@�C�������݂���ꍇ�͉������Ȃ�
            if (File.Exists(backupPath))
                return;

            // �t�@�C�����R�s�[���ăo�b�N�A�b�v�쐬
            File.Copy(originalPath, backupPath);
        }
    }
    catch (Exception ex)
    {
        MessageBox.Show($"�o�b�N�A�b�v�ۑ����ɃG���[���������܂���: {ex.Message}", "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}
```
