// Utils.ProcessImageMarkers.cs

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    internal partial class Utils
    {
        /// <summary>
        /// webhelp�f�B���N�g�����̂��ׂĂ�HTML�t�@�C�����������AIMAGEMARKER�Ɋ�Â��ĉ摜��src��ύX
        /// </summary>
        /// <param name="webhelpDirectory">webhelp�f�B���N�g���̃p�X</param>
        /// <param name="extractedImagesDirectory">extracted_images�f�B���N�g���̑��΃p�X�i��: "extracted_images"�j</param>
        /// <returns>�������ꂽ�t�@�C����</returns>
        public static int ProcessImageMarkersInWebhelp(
            string webhelpDirectory,
            string extractedImagesDirectory = "extracted_images")
        {
            if (string.IsNullOrEmpty(webhelpDirectory))
                throw new ArgumentException("webhelp�f�B���N�g�����w�肳��Ă��܂���B", nameof(webhelpDirectory));

            if (!Directory.Exists(webhelpDirectory))
                throw new DirectoryNotFoundException($"�w�肳�ꂽ�f�B���N�g�������݂��܂���: {webhelpDirectory}");

            int processedFileCount = 0;

            try
            {
                // webhelp�f�B���N�g�����̂��ׂĂ�HTML�t�@�C�����擾
                var htmlFiles = Directory.GetFiles(webhelpDirectory, "*.html", SearchOption.AllDirectories);

                foreach (string htmlFilePath in htmlFiles)
                {
                    try
                    {
                        if (ProcessSingleHtmlFile(htmlFilePath, extractedImagesDirectory))
                        {
                            processedFileCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"�t�@�C�������G���[ ({htmlFilePath}): {ex.Message}");
                    }
                }

                return processedFileCount;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"HTML�摜�}�[�J�[�������ɃG���[���������܂���: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// �P���HTML�t�@�C�����������AIMAGEMARKER�Ɋ�Â��ĉ摜��src��ύX
        /// </summary>
        /// <param name="htmlFilePath">�����Ώۂ�HTML�t�@�C���p�X</param>
        /// <param name="extractedImagesDirectory">extracted_images�f�B���N�g���̑��΃p�X</param>
        /// <returns>�ύX���s��ꂽ�ꍇ��true</returns>
        private static bool ProcessSingleHtmlFile(string htmlFilePath, string extractedImagesDirectory)
        {
            // HTML�t�@�C����ǂݍ���
            string htmlContent;
            using (var reader = new StreamReader(htmlFilePath, Encoding.UTF8))
            {
                htmlContent = reader.ReadToEnd();
            }

            string originalContent = htmlContent;

            // �摜�ƃ}�[�J�[�̃y�A������
            htmlContent = ProcessImageAndMarkerPairs(htmlContent, extractedImagesDirectory);

            // �c���IMAGEMARKER���폜
            htmlContent = RemoveRemainingImageMarkers(htmlContent);

            // �ύX���������ꍇ�̂݃t�@�C����ۑ�
            if (htmlContent != originalContent)
            {
                using (var writer = new StreamWriter(htmlFilePath, false, Encoding.UTF8))
                {
                    writer.Write(htmlContent);
                }
                return true;
            }

            return false;
        }

        /// <summary>
        /// <img>�^�O�̒���ɂ���[IMAGEMARKER:xxx]���������Aimg��src��ύX
        /// </summary>
        /// <param name="htmlContent">�����Ώۂ�HTML���e</param>
        /// <param name="extractedImagesDirectory">extracted_images�f�B���N�g���̑��΃p�X</param>
        /// <returns>�������ꂽHTML���e</returns>
        private static string ProcessImageAndMarkerPairs(string htmlContent, string extractedImagesDirectory)
        {
            // <img>�^�O�̒����<p>�^�O�ň͂܂ꂽ[IMAGEMARKER:xxx]������p�^�[��������
            // �p�^�[������:
            // (<img[^>]*>) - <img>�^�O���L���v�`��
            // \s*</p>\s* - </p>�^�O�Ƃ��̑O��̋�
            // <p[^>]*>\s* - <p>�^�O�̊J�n�Ƃ��̌�̋�
            // \[IMAGEMARKER:([^\]]+)\] - [IMAGEMARKER:xxx]�̌`���ŁAxxx�̕������L���v�`��
            // \s*</p> - ���̌�̋󔒂�</p>�^�O
            string pattern = @"(<img[^>]*>)\s*</p>\s*<p[^>]*>\s*\[IMAGEMARKER:([^\]]+)\]\s*</p>";

            var regex = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);

            htmlContent = regex.Replace(htmlContent, match =>
            {
                string imgTag = match.Groups[1].Value;
                string markerValue = match.Groups[2].Value;

                // img�^�O��src������V�����p�X�ɕύX
                string newSrc = $"{extractedImagesDirectory}/{markerValue}.png";
                string updatedImgTag = UpdateImageSrc(imgTag, newSrc);

                // �}�[�J�[���폜���A�X�V���ꂽimg�^�O�݂̂�Ԃ�
                return updatedImgTag + "</p>";
            });

            return htmlContent;
        }

        /// <summary>
        /// <img>�^�O��src������V�����l�ɕύX
        /// </summary>
        /// <param name="imgTag">����<img>�^�O</param>
        /// <param name="newSrc">�V����src�l</param>
        /// <returns>src�������X�V���ꂽ<img>�^�O</returns>
        private static string UpdateImageSrc(string imgTag, string newSrc)
        {
            // src�����̃p�^�[�����}�b�`
            var srcPattern = @"src\s*=\s*[""']([^""']*)[""']";
            var srcRegex = new Regex(srcPattern, RegexOptions.IgnoreCase);

            if (srcRegex.IsMatch(imgTag))
            {
                // ������src������V�����l�ɒu��
                return srcRegex.Replace(imgTag, $"src=\"{newSrc}\"");
            }
            else
            {
                // src���������݂��Ȃ��ꍇ�͒ǉ�
                // <img �̒���ɑ}��
                var insertPattern = @"(<img)(\s|>)";
                var insertRegex = new Regex(insertPattern, RegexOptions.IgnoreCase);
                
                if (insertRegex.IsMatch(imgTag))
                {
                    return insertRegex.Replace(imgTag, $"$1 src=\"{newSrc}\"$2");
                }
            }

            return imgTag;
        }

        /// <summary>
        /// �c���[IMAGEMARKER:xxx]�p�^�[�������ׂč폜
        /// </summary>
        /// <param name="htmlContent">�����Ώۂ�HTML���e</param>
        /// <returns>�}�[�J�[���폜���ꂽHTML���e</returns>
        private static string RemoveRemainingImageMarkers(string htmlContent)
        {
            // <p>�^�O�ň͂܂ꂽ[IMAGEMARKER:xxx]�p�^�[�����폜
            string paragraphMarkerPattern = @"<p[^>]*>\s*\[IMAGEMARKER:[^\]]+\]\s*</p>";
            htmlContent = Regex.Replace(htmlContent, paragraphMarkerPattern, "", RegexOptions.IgnoreCase | RegexOptions.Multiline);

            // ���̑��̏ꏊ�ɂ���[IMAGEMARKER:xxx]�p�^�[�����폜
            string markerPattern = @"\[IMAGEMARKER:[^\]]+\]";
            htmlContent = Regex.Replace(htmlContent, markerPattern, "", RegexOptions.IgnoreCase);

            // �A�������s��]���ȋ󔒂𐮗�
            htmlContent = Regex.Replace(htmlContent, @"\n\s*\n\s*\n", "\n\n");

            return htmlContent;
        }

        /// <summary>
        /// �������v�����擾
        /// </summary>
        /// <param name="webhelpDirectory">webhelp�f�B���N�g���̃p�X</param>
        /// <returns>���v���̕�����</returns>
        public static string GetImageMarkerProcessingStatistics(string webhelpDirectory)
        {
            if (string.IsNullOrEmpty(webhelpDirectory) || !Directory.Exists(webhelpDirectory))
                return "�w�肳�ꂽ�f�B���N�g�������݂��܂���B";

            try
            {
                var htmlFiles = Directory.GetFiles(webhelpDirectory, "*.html", SearchOption.AllDirectories);
                int totalFiles = htmlFiles.Length;
                int filesWithMarkers = 0;
                int totalMarkers = 0;

                foreach (string htmlFilePath in htmlFiles)
                {
                    try
                    {
                        string content;
                        using (var reader = new StreamReader(htmlFilePath, Encoding.UTF8))
                        {
                            content = reader.ReadToEnd();
                        }

                        var markerMatches = Regex.Matches(content, @"\[IMAGEMARKER:[^\]]+\]", RegexOptions.IgnoreCase);
                        if (markerMatches.Count > 0)
                        {
                            filesWithMarkers++;
                            totalMarkers += markerMatches.Count;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"���v�擾�G���[ ({htmlFilePath}): {ex.Message}");
                    }
                }

                var statistics = new StringBuilder();
                statistics.AppendLine("�摜�}�[�J�[�������v:");
                statistics.AppendLine("====================");
                statistics.AppendLine($"��HTML�t�@�C����: {totalFiles}");
                statistics.AppendLine($"�}�[�J�[���܂ރt�@�C����: {filesWithMarkers}");
                statistics.AppendLine($"���}�[�J�[��: {totalMarkers}");

                return statistics.ToString();
            }
            catch (Exception ex)
            {
                return $"���v�擾���ɃG���[���������܂���: {ex.Message}";
            }
        }

        /// <summary>
        /// extracted_images�f�B���N�g�����̃t�@�C���ꗗ���擾
        /// </summary>
        /// <param name="webhelpDirectory">webhelp�f�B���N�g���̃p�X</param>
        /// <param name="extractedImagesDirectory">extracted_images�f�B���N�g���̑��΃p�X</param>
        /// <returns>���o���ꂽ�摜�t�@�C���̃p�X�ꗗ</returns>
        public static List<string> GetExtractedImageFiles(
            string webhelpDirectory,
            string extractedImagesDirectory = "extracted_images")
        {
            string extractedImagesPath = Path.Combine(webhelpDirectory, extractedImagesDirectory);
            
            if (!Directory.Exists(extractedImagesPath))
                return new List<string>();

            try
            {
                return Directory.GetFiles(extractedImagesPath, "*.png", SearchOption.TopDirectoryOnly)
                    .Select(Path.GetFileName)
                    .ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"�摜�t�@�C���ꗗ�擾�G���[: {ex.Message}");
                return new List<string>();
            }
        }
    }
}