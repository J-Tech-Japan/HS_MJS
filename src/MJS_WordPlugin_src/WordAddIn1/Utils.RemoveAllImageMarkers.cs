// Utils.RemoveAllImageMarkers.cs

using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    internal partial class Utils
    {

        /// <summary>
        /// Word�h�L�������g����ExtractImagesAndCanvasFromWordWithText�Őݒ肳�ꂽ�S�Ẳ摜�}�[�J�[���폜����
        /// </summary>
        /// <param name="document">�}�[�J�[���폜����Ώۂ�Word�h�L�������g</param>
        /// <returns>�폜���ꂽ�}�[�J�[�̐�</returns>
        public static int RemoveAllImageMarkers(Word.Document document)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            int removedCount = 0;

            try
            {
                // �h�L�������g�S�͈̂̔͂��擾
                var fullRange = document.Range();

                // ����������ݒ�
                var find = fullRange.Find;
                find.ClearFormatting();
                find.Text = @"\[IMAGEMARKER:*\]";
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = true;
                find.MatchSoundsLike = false;
                find.MatchAllWordForms = false;

                // �}�[�J�[�������������č폜
                while (find.Execute())
                {
                    try
                    {
                        // �}�[�J�[�e�L�X�g�͈̔͂��擾
                        var markerRange = fullRange.Duplicate;

                        // �}�[�J�[�̑O��̉��s���܂߂č폜�͈͂��g��
                        ExtendRangeToIncludeAssociatedLineBreaks(markerRange);

                        // �}�[�J�[�Ƃ��̑O��̉��s���폜
                        markerRange.Delete();
                        removedCount++;

                        // �폜��A�����͈͂����Z�b�g
                        fullRange.SetRange(0, document.Range().End);
                        find.ClearFormatting();
                        find.Text = @"\[IMAGEMARKER:*\]";
                        find.MatchWildcards = true;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"�ʃ}�[�J�[�폜�G���[: {ex.Message}");
                        // �ʂ̃}�[�J�[�폜�ŃG���[���������Ă��������p��
                        break;
                    }
                }

                return removedCount;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"�摜�}�[�J�[�폜���ɃG���[���������܂���: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// �}�[�J�[�͈͂��g�����āA�֘A������s���܂߂�悤�ɒ���
        /// </summary>
        /// <param name="markerRange">�}�[�J�[�͈̔�</param>
        private static void ExtendRangeToIncludeAssociatedLineBreaks(Word.Range markerRange)
        {
            try
            {
                // �}�[�J�[�̑O�̕������`�F�b�N
                if (markerRange.Start > 0)
                {
                    var beforeRange = markerRange.Document.Range(markerRange.Start - 1, markerRange.Start);
                    if (beforeRange.Text == "\r" || beforeRange.Text == "\n")
                    {
                        // �O�̉��s���폜�͈͂Ɋ܂߂�
                        markerRange.SetRange(markerRange.Start - 1, markerRange.End);
                    }
                }

                // �}�[�J�[�̌�̕������`�F�b�N
                if (markerRange.End < markerRange.Document.Range().End)
                {
                    var afterRange = markerRange.Document.Range(markerRange.End, markerRange.End + 1);
                    if (afterRange.Text == "\r" || afterRange.Text == "\n")
                    {
                        // ��̉��s���폜�͈͂Ɋ܂߂�
                        markerRange.SetRange(markerRange.Start, markerRange.End + 1);

                        // �A��������s���`�F�b�N
                        if (markerRange.End < markerRange.Document.Range().End)
                        {
                            var nextAfterRange = markerRange.Document.Range(markerRange.End, markerRange.End + 1);
                            if (nextAfterRange.Text == "\r" || nextAfterRange.Text == "\n")
                            {
                                markerRange.SetRange(markerRange.Start, markerRange.End + 1);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"�͈͊g���G���[: {ex.Message}");
            }
        }

        /// <summary>
        /// ����̉摜�}�[�J�[�݂̂��폜����
        /// </summary>
        /// <param name="document">�}�[�J�[���폜����Ώۂ�Word�h�L�������g</param>
        /// <param name="markerText">�폜�������}�[�J�[�̃e�L�X�g�i�g���q�Ȃ��̃t�@�C�����j</param>
        /// <returns>�폜���ꂽ�}�[�J�[�������������ǂ���</returns>
        public static bool RemoveSpecificImageMarker(Word.Document document, string markerText)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            if (string.IsNullOrEmpty(markerText))
                throw new ArgumentException("�}�[�J�[�e�L�X�g���w�肳��Ă��܂���B", nameof(markerText));

            try
            {
                // �h�L�������g�S�͈̂̔͂��擾
                var fullRange = document.Range();

                // ����������ݒ�
                var find = fullRange.Find;
                find.ClearFormatting();
                find.Text = $"[IMAGEMARKER:{markerText}]";
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = false;

                // �}�[�J�[������
                if (find.Execute())
                {
                    // �}�[�J�[�e�L�X�g�͈̔͂��擾
                    var markerRange = fullRange.Duplicate;

                    // �}�[�J�[�̑O��̉��s���܂߂č폜�͈͂��g��
                    ExtendRangeToIncludeAssociatedLineBreaks(markerRange);

                    // �}�[�J�[�Ƃ��̑O��̉��s���폜
                    markerRange.Delete();

                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"�摜�}�[�J�[�폜���ɃG���[���������܂���: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// �h�L�������g���̉摜�}�[�J�[�̐����擾����
        /// </summary>
        /// <param name="document">�Ώۂ�Word�h�L�������g</param>
        /// <returns>���������}�[�J�[�̐�</returns>
        public static int CountImageMarkers(Word.Document document)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            int count = 0;

            try
            {
                // �h�L�������g�S�͈̂̔͂��擾
                var fullRange = document.Range();

                // ����������ݒ�
                var find = fullRange.Find;
                find.ClearFormatting();
                find.Text = @"\[IMAGEMARKER:*\]";
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = true;

                // �}�[�J�[�������������ăJ�E���g
                while (find.Execute())
                {
                    count++;

                    // ���̌����̂��߂ɔ͈͂𒲐�
                    fullRange.SetRange(fullRange.End, document.Range().End);
                    if (fullRange.Start >= fullRange.End)
                        break;
                }

                return count;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"�摜�}�[�J�[�J�E���g�G���[: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// �h�L�������g���̑S�Ẳ摜�}�[�J�[�̈ꗗ���擾����
        /// </summary>
        /// <param name="document">�Ώۂ�Word�h�L�������g</param>
        /// <returns>���������}�[�J�[�e�L�X�g�̃��X�g</returns>
        public static List<string> GetImageMarkersList(Word.Document document)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            var markers = new List<string>();

            try
            {
                // �h�L�������g�S�͈̂̔͂��擾
                var fullRange = document.Range();

                // ����������ݒ�
                var find = fullRange.Find;
                find.ClearFormatting();
                find.Text = @"\[IMAGEMARKER:*\]";
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = true;

                // �}�[�J�[�������������Ĉꗗ�ɒǉ�
                while (find.Execute())
                {
                    try
                    {
                        string markerText = fullRange.Text;
                        markers.Add(markerText);

                        // ���̌����̂��߂ɔ͈͂𒲐�
                        fullRange.SetRange(fullRange.End, document.Range().End);
                        if (fullRange.Start >= fullRange.End)
                            break;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"�ʃ}�[�J�[�擾�G���[: {ex.Message}");
                        break;
                    }
                }

                return markers;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"�摜�}�[�J�[�ꗗ�擾�G���[: {ex.Message}");
                return new List<string>();
            }
        }
    }
}