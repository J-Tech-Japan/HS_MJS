using System.Management.Automation;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.PowerShell;

// �A�Z���u���Ɋւ����ʏ��́A�ȉ��̑����Z�b�g�ɂ����
// ���䂳��܂��B�A�Z���u���Ɋ֘A�t�����Ă������ύX����ɂ́A
// �����̑����l��ύX���܂��B
[assembly: AssemblyTitle("MJS���[�h�v���O�C��")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("MJS���[�h�v���O�C��")]
[assembly: AssemblyCopyright("Copyright c  2017")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// ComVisible �� false �ɐݒ肷��ƁA���̌^�͂��̃A�Z���u������ COM �R���|�[�l���g����
// �Q�ƕs�\�ɂȂ�܂��BCOM ���炱�̃A�Z���u�����̌^�ɃA�N�Z�X����ꍇ�́A
// ���̌^�� ComVisible ������ true �ɐݒ肵�Ă��������B
[assembly: ComVisible(false)]

// ���̃v���W�F�N�g�� COM �Ɍ��J�����ꍇ�A���� GUID �� typelib �� ID �ɂȂ�܂�
[assembly: Guid("efcb7755-f1d8-4bb1-b051-137af1a308da")]

// Word�̃��{���ɂ̓��r�W����������3�̐����Ńo�[�W������\�����܂�
// [���W���[ �o�[�W����.�}�C�i�[ �o�[�W����.�r���h�ԍ�]
// �����l�� 3.1.0 �ł��B
// �ȉ��̐����̓r���h����x�Ɏ����I�ɃC���N�������g����܂��B
// (PowerShell�X�N���v�g IncrementMinorVersion.ps1 �ŊǗ�����܂��B
// �����̐����̓��r�W�����ԍ��ł��BWord�ɂ͕\������܂���B
[assembly: AssemblyVersion("3.1.7.0")]
[assembly: AssemblyFileVersion("3.1.7.0")]

// �ʏ�̓r���h�O�C�x���g�R�}���h���C���Ɉȉ��̃R�}���h���ݒ肳��Ă��܂��i�����[�X�r���h���ɃC���N�������g�j�B
// if "$(ConfigurationName)"=="Release" powershell -ExecutionPolicy Bypass -File "$(ProjectDir)IncrementMinorVersion.ps1"

// �f�o�b�O�r���h�ł��C���N�������g�������ꍇ�́A�ȉ��̃R�[�h�ɏ��������܂��B
// powershell -ExecutionPolicy Bypass -File "$(ProjectDir)IncrementMinorVersion.ps1"

// �o�[�W������3.1.0�Ƀ��Z�b�g����ꍇ�́A�ȉ��̂悤�ɏ��������܂��B
// powershell - ExecutionPolicy Bypass - File "$(ProjectDir)IncrementMinorVersion.ps1" reset
