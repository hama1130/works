<#
�g�p���@�F./export-VBA.ps1 .\sample.xlsm
        �i�G�C���A�X�ݒ�𐄏��j
�@�\�@�@�F�����Ŏw�肵���t�@�C������VBA���e�L�X�g�t�@�C���Ƃ��ďo�͂���B
���l�@�@�F���O��Excel�̃I�v�V��������uVBA�v���W�F�N�g�I�u�W�F�N�g���f���ւ̃A�N�Z�X��M������v�̐ݒ肪�K�v
#>

#�����`�F�b�N
if($args.Count -ne 1){
    echo "�����Ƀt�@�C�������w�肵�Ă�������"
    exit
}

#�g���q�`�F�b�N
$filename = $args[0]
if(-not ($filename -like "*.xlsm")){
    echo ".xlsm�t�@�C�����w�肵�Ă�������"
    exit
}

#��΃p�X�ɕύX
$filename = $(resolve-path $filename).ProviderPath

#VBObject�̎�ވꗗ
$hash = @{
    1=".bas";           #1 : �W�����W���[��(bas)
    2=".cls";           #2: �N���X���W���[��(cls)
    3=".frm";           #3: �t�H�[��(.frm)
    11=".cls";          #11: ActivezX (.cls)
    100=".cls"          #100: �h�L�������g�E�V�[�g(.cls)
}

#Excel�A�v�����J��
$exl =new-object -comobject excel.application
# $exl.visible =$true 
$exl.DisplayAlerts =$False
$exl.EnableEvents = $False

#�u�b�N���J��(�ǂݎ���p�j�@���R�Ԗڂ�ReadOnly����
$wb=$exl.Workbooks.Open($filename, $null, $true)

#�o�͐�t�H���_���쐬
$dir_name = "VBA_$($wb.name)".replace(".xlsm","")
mkdir $dir_name | cd

#�G�N�X�|�[�g
$vbcmps = $wb.VBProject.VBComponents
 foreach ($tmp in $vbcmps){
    $tmp.Export($(pwd).ProviderPath + "\"+$tmp.name + $hash[$tmp.Type])
    echo "$($tmp.name)���G�N�X�|�[�g"
 }

#Excel�����(�ۑ����Ȃ��j
$wb.close($False)
$exl.EnableEvents = $True
$exl.Quit()

#�I������
$exl = $null
[GC]::Collect()

#���b�Z�[�W
echo "�G�N�X�|�[�g���I�����܂���"

#���̃t�H���_�ɖ߂�
cd ..