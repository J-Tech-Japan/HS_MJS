# MJS Word�v���O�C���̊T�v
MJS Word�v���O�C���́AWord���e�̕ҏW�x���A�����HTML�o�͂̂��߂̃c�[���ł��B

## �A�[�L�e�N�`��
- VSTO�iVisual Studio Tools for Office�j�x�[�X��Word�A�h�C��
- .NET Framework 4.8 / C# 7.3
- COM Interop���g�p����Office�A�g
- XML/HTML�����ɂ�镶���ϊ�
- ���{��UI �ɂ�鑀��C���^�[�t�F�[�X

## �v���O�C���̋@�\
Word�v���O�C�����C���X�g�[������Ă����Ԃ�Word���N������ƁA���j���[�ɁuMJS���[�h�v���O�C���v�^�u���ǉ�����܂��B
�uMJS���[�h�v���O�C���v���N���b�N����ƁA�����珇�Ɂu�������o�́v�A�u�����N�ݒ�v�A�u�X�^�C���`�F�b�N�v�A�uHTML�o�́v�̃��{�����\������܂��B

### �������o��
- Word���e�̖ڎ��\���̏��ł���u�������v���A�e�L�X�g�t�@�C���i�������t�@�C���ƌĂт܂��j�Ƃ��ďo�́i�X�V�j���܂��B
- �������t�@�C���́AWord���e�idoc�j�Ɠ����K�w�ɂ���uheaderFile�v�t�H���_���ɕۑ�����܂��B

#### �������o�͂Ɋ֘A����t�@�C��
- BookInfoButton.cs: �������o�̓{�^���̏����BmakeBookInfo()���\�b�h���Ăяo���ă��C�����������s
- RibbonMJS.MakeBookInfo.cs: �������쐬�̃��C�������B�������̌��o���i������͂��ău�b�N�}�[�N�𐶐����A�V���������̔�r�E�X�V�����s
- RibbonMJS.MakeBookInfo.Helper.cs: �������쐬�̃w���p�[�֐��Q�B�t�@�C�����`�F�b�N�A�u�b�N�}�[�N����A�V���������̔�r����
- RibbonMJS.CheckDocInfo.cs: �V���������̔�r�����B���ڂ̒ǉ��E�폜�E�ύX�EID�s��v�E�^�C�g���ύX�Ȃ�8��ނ̕ύX�p�^�[�������o���A��r���ʃ��X�g�𐶐�
- RibbonMJS.CheckSortInfo.cs: ��������r���ʂ̃\�[�g�����B���ԊK�w�i4���x���j�Ɋ�Â��K�؂ȕ��я��ł̔�r���ʕ\���@�\
- RibbonMJS.MakeBookInfo.HeaderFile.cs: �w�b�_�[�t�@�C���i�������t�@�C���j�̓ǂݍ��݁E�������ݏ����ƃt�@�C���A�N�Z�X����
- BookInfo.cs: �������̃f�t�H���g�l���͗p�_�C�A���O�t�H�[���B2���̐��l���͂ƑS�p�E���p�ϊ��@�\
- HeadingInfo.cs: ���o�������i�[����f�[�^�N���X�i���ԁE�^�C�g���EID�E�}�[�W����j

### �����N�ݒ�
- Word���e�̕ҏW�⏕�@�\�Ƃ��āA���̍��ڂւ̃����N�i�Q�Ɓj��ݒ肵�܂��B
- �������t�@�C�����獀�ڂ�ǂݍ��݁A���΃p�X�v�Z��URL�������s���ăn�C�p�[�����N���쐬���܂��B

#### �����N�ݒ�Ɋ֘A����t�@�C��
- SetLink.cs: �����N�ݒ�_�C�A���O�t�H�[��
- SetLinkButton.cs: �����N�ݒ�{�^���̏����BSetLink�_�C�A���O�t�H�[����\�����ă����N�ݒ�@�\���N��

### �X�^�C���`�F�b�N
- HTML�o�͂̑O�����Ƃ��āAWord���e���ŋK��ȊO�̃X�^�C�����g���Ă��Ȃ������`�F�b�N���܂��B

#### �X�^�C���`�F�b�N�Ɋ֘A����t�@�C��
- StyleCheckButton.cs: �X�^�C���`�F�b�N�̃��C�������B�e���v���[�g���狖���ꂽMJS�X�^�C�����X�g���擾���A�h�L�������g�S�̂̌��؂����s
- StyleCheckButton.HandleProcess.cs: �X�^�C���`�F�b�N������̌��ʏ����i�����E���s�E��~���̃��b�Z�[�W�\���ƃ{�^������j
- StyleCheckButton.Paragraphs.cs: �e�i���̃X�^�C�����؏����BMJS�X�^�C���K�����`�F�b�N�Ǝ菇�ԍ����Z�b�g�p�X�^�C���̐���������
- StyleCheckButton.NonInlineShape.cs: �}�`�E�摜�̔z�u�`�F�b�N�����B�s���z�u�ȊO�̃V�F�C�v��`��L�����o�X�̔z�u�G���[�����o

### HTML�o��
- Word���e�̓��e���AHTML�ŏo�͂��܂��B
- HTML�t�@�C���͂��ׂ�webhelp�t�H���_�ɕۑ�����܂��B
- Word���e����擾�����摜�́Awebhelp�t�H���_����pict�t�H���_�ɕۑ�����܂��B

#### HTML�o�͂Ɋ֘A����t�@�C��
- GenerateHTMLButton.cs: HTML�o�͂̃��C�������BWord�h�L�������g����Web�w���v�`����HTML�R���e���c�𐶐����A�\���摜�̒��o���猟���@�\�܂ŕ�I�ȕϊ����������s
- GenerateHTMLButton.CopyDocumentToHtml.cs: Word�h�L�������g��HTML�ϊ��p�ɕ������鏈���B�N���b�v�{�[�h�o�R�Ńh�L�������g�S�̂��R�s�[���A�V�K�h�L�������g�ɓ\��t��
- GenerateHTMLButton.StyleProcessor.cs: Word�������璊�o����CSS�X�^�C����`�̉�͏����Bmso-style-name��������ɏ͕����N���X��X�^�C���������𐶐�
- GenerateHTMLButton.ProcessHTML.cs: �ꎞ�I�ɕۑ����ꂽHTML�t�@�C���̓ǂݍ��݂ƑO�����B�����G���R�[�f�B���O�C����HTML�\���̐��K�������s
- GenerateHTMLButton.HtmlTemplate1.cs: �ʃy�[�W�pHTML�e���v���[�g�����B�p���������X�g�A�ڎ��K�w�A�����@�\���܂ޕW���y�[�W���C�A�E�g�̍\�z
- GenerateHTMLButton.HtmlCoverTemplate.cs: �\���y�[�W�pHTML�e���v���[�g�����B���i���S�A�^�C�g���A���W�����܂ޕ\�����C�A�E�g�̍\�z
- GenerateHTMLButton.IdxHtmlTemplate.cs: �C���f�b�N�X�i�ڎ��j�y�[�W�pHTML�e���v���[�g�����B�S�̂̃i�r�Q�[�V�����\���ƃt���[���ݒ���܂ރ��C���y�[�W�̍\�z
- GenerateHTMLButton.CollectInfo.cs: Word��������\���E���W�E�o�[�W�������̎��W�����B����X�^�C���̒i������^�C�g���⒘�쌠���𒊏o
- GenerateHTMLButton.CollectMergeScript.cs: ���o�����������������t�@�C���iheaderFile�j������W���AHTML�o�͎��̃y�[�W�}�[�W�����p�����𐶐�
- GenerateHTMLButton.Helper.cs: HTML�����ŋ��ʗ��p����w���p�[�֐��Q�B�p�X�����A�t�@�C������A�\���I���_�C�A���O�A��O�����Ȃ�
- GenerateHTMLButton.CopyImagesFromAppDataLocalTemp.cs: AppData/Local/Temp�t�H���_����摜�t�@�C�����������Awebhelp/pict�t�H���_�ɓK�؂ȃt�@�C�����ŃR�s�[
- RibbonMJS.InnerNode.cs: HTML�ϊ�����XML�m�[�h���������BWord�����̊e�v�f�i�\�E�}�`�E�X�^�C���j��HTML�v�f�ɕϊ����郁�C������
- RibbonMJS.InnerNode.Helper.cs: InnerNode.cs�̃w���p�[�֐��Q�B����菇�EQ&A�E�I�����E�ӏ������E�\�E�R�����ȂǊe��MJS�X�^�C���̐�pHTML�ϊ�����


#### XML�EHTML�ϊ��֘A�̃t�@�C��
- GenerateHTMLButton.XMLProcessDocument.cs: Word��HTML�o�͂�XML�`���ɕϊ����A�ڎ��E�{���\���̉�͂ƕ������������s
- GenerateHTMLButton.XMLBuildTocBody.cs: XML�`���̕����f�[�^����ڎ��\���Ɩ{���y�[�W���\�z�B�͕�����y�[�W�K�w�̐�������
- GenerateHTMLButton.XMLExportTocAsJsFiles.cs: �ڎ��f�[�^��JavaScript�`���ŏo�͂��AWeb�w���v�V�X�e���̃i�r�Q�[�V�����@�\�𐶐�

#### �����@�\�֘A�̃t�@�C��
- GenerateHTMLButton.SearchIndex.cs: �����ΏۂƂȂ�HTML�y�[�W�ƌ����p�C���f�b�N�X�t�@�C���isearch.js�j�𐶐����A������b�̒��o�ƍ����������s
- GenerateHTMLButton.RemoveSearchBlock.cs: �w�肳�ꂽ�^�C�g���̃y�[�W���猟���u���b�N���폜���A�����ΏۊO�R���e���c�̏��O���������s

## ���̑��̃t�@�C��

### ���ʃ��[�e�B���e�B�N���X�iUtils�j
- Utils.FileIO.cs: ���X�g�^�ϐ��̓��e���e�L�X�g�t�@�C���ɏ������ޔėp���\�b�h�ƁA�t�@�C������̋��ʏ������
- Utils.TextProcessing.cs: �S�p�������甼�p�����ւ̕ϊ��@�\�B�����E�p���E�L���̕�����ϊ�������ConvertWideToNarrow���\�b�h���
- Utils.RemoveSpanTagFromHtml.cs: HTML�t�@�C������s�v��span�^�O���폜����@�\�BHTML�t�@�C���P�̂܂��̓t�H���_���̈ꊇ�����ɑΉ����A�����Ȃ��̃V���v����span�^�O�݂̂�ΏۂƂ��Ē��g�̃e�L�X�g�͕ێ�
- Utils.ExtractImagesFromWord.cs: Word�h�L�������g����EnhMetaFileBits���g�p���ăC�����C���}�`�E�t���[�e�B���O�}�`�E�L�����o�X�}�`�����i���Œ��o���鏈���B���o�����摜�ɑΉ�����}�[�J�[��Word�������ɑ}�����A�㑱��HTML�����ŉ摜�p�X�𐳊m�ɎQ�Ƃł���悤����
- Utils.ExtractImagesFromWord.CheckStyle.cs: �摜���o���̃X�^�C�����菈���BMJS����X�^�C���i�摜�A�菇���A�{�����A�R�������A�\���A�����t���[���j�̔���ƃX�^�C���x�[�X�������o�E�X�L�b�v����A�\���Z�N�V��������@�\���
- Utils.ExtractImagesFromWord.Info.cs: �摜���o���ʂ̓��v��񐶐��ƃe�L�X�g�t�@�C���o�͋@�\�B���o�����摜�̎�ʁE�T�C�Y�E���ʂɊւ���ڍ׃��|�[�g�𐶐�
- Utils.ExtractImagesFromWord.InsertMarker.cs: ���o�����摜�̈ʒu�Ƀ}�[�J�[�e�L�X�g��}�����鏈���B�C�����C���}�`�ƃt���[�e�B���O�}�`���ꂼ��ɑΉ�����[IMAGEMARKER:xxx]�`���̃}�[�J�[�}���@�\
- Utils.ProcessImageMarkers.cs: HTML�o�͌��webhelp�f�B���N�g�����ŁA[IMAGEMARKER:xxx]�p�^�[�����������Ή�����摜�t�@�C���ւ�src�������X�V���鏈���B�摜�}�[�J�[��HTML��img�v�f���֘A�t���āA���o���ꂽ�摜�ւ̐��m�ȃ����N�𐶐�
- Utils.RemoveAllImageMarkers.cs: Word��������摜�}�[�J�[�e�L�X�g���폜���鏈���BHTML�o�͊�����̃N���[���A�b�v�@�\
- Utils.RemoveImageMarkersFromSearchJs.cs: �����@�\�pJavaScript�isearch.js�j�t�@�C������摜�}�[�J�[�e�L�X�g���폜���A�����ΏۊO�R���e���c�Ƃ��ď��O���鏈��

### �ݒ�E�������֘A�̃t�@�C��
- RibbonMJS.Config.cs: HTML�o�͗p�p�X�ꗗ�̏����A�e��萔�E�p�^�[���̒�`�A���������ݒ�Ȃǃv���O�C���S�̂̐ݒ�@�\���

### UI�E�V�X�e���@�\
- RibbonMJS.ClearClipboard.cs: �N���b�v�{�[�h�̈��S�ȃN���A�����BCOMException�Ή��̃��g���C�@�\�t���N���b�v�{�[�h����
- RibbonMJS.Designer.cs: ���{��UI�iMJS���[�h�v���O�C���^�u�j�̃f�U�C�i�[���������R�[�h�B�{�^���z�u�E�C�x���g�n���h���[�ݒ�E���\�[�X�Ǘ�


## Copilot ���t�@�N�^�����O�w����

### ��ʓI�ȃK�C�h���C��
- �w�肪�Ȃ�����A�����̋@�\���ێ������܂܃R�[�h��ύX���Ă��������B
- �ϐ����⃁�\�b�h���́A�����̓��e���c���ł���悤�ɋL�q���Ă��������B
- �l�X�g�����炷���߁A�������^�[���iearly return�j��S�����Ă��������B
- �}�W�b�N�i���o�[�͔����A�萔��enum���g�p���Ă��������B
- �K�؂ɗ�O�������s���A�K�v�ɉ����ăG���[���O���o�͂��Ă��������B
- ����Ń��O�̋L�^���K�v���Ǝv�����ӏ��ɂ́A�K�؂ȃ��O�o�͂�ǉ����Ă��������B

### C#/.NET �ŗL
- �E�ӂ���^�����m�ȏꍇ�̂� `var` ���g�p���Ă��������B
- `IDisposable` �ȃI�u�W�F�N�g�� `using` ���Ŋm���ɔj�����Ă��������B
- �R���N�V������I�u�W�F�N�g�������q��ϋɓI�ɗ��p���Ă��������B
- ������A���ɂ͉\�Ȍ��蕶�����ԁi$"...") ���g���Ă��������B
- LINQ�Ȃǂ̋@�\�����p���A�ǐ��̍����R�[�h�������Ă��������B

### �A�h�C���^Interop �ŗL
- ���������[�N�h�~�̂��߁ACOM�I�u�W�F�N�g�͊m���ɉ�����Ă��������B
- Interop�̃��W�b�N�̓w���p�[���\�b�h��N���X�ɂ܂Ƃ߂Ă��������B
