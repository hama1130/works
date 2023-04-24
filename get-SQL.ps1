<#
�g�p���@�F./get-SQL.ps1 "DB.dbo.testprocedure"
�T�v�@�@�FDWH�T�[�o�[�ɐڑ����A�I�u�W�F�N�g�̒�`�����R���\�[���ɏo�͂���
�����@�@�F�I�u�W�F�N�g���͈�������^����B
�@�@�@�@�@���_�C���N�g�ɂ��e�L�X�g�t�@�C���ւ̏o�͂��\�B
�@�@�@�@�@SQL Server Management Studio����쐬�����X�N���v�g�Ɠ��l�̌`���ɒ����ς݁B�@
�����@�@�@�F�I�u�W�F�N�g��
#>

# ������`�i�I�u�W�F�N�g���j
Param([Parameter(Mandatory = $true)][String]$object_name)

#�����𕪉�
$ary = $object_name.split(".")
$dbname = $ary[0]
$schema = $ary[1]
$name = $ary[2]

#�_�u���N�I�[�e�[�V����������ƃG���[�ɂȂ�̂ŏ���
$object_name = $object_name.Replace('"','')

#�萔
$datasource = "localhost" #SQL�T�[�o�[�̃z�X�g
$database_name = $dbname #�f�[�^�x�[�X��
$SQLstring = "select definition from $dbname.sys.sql_modules where object_id=object_id('$object_name')" #SQL������1
$SQLstring2 = "select type from $dbname.sys.objects where object_id=object_id('$object_name')" #SQL������2

# SqlConnectionStringBuilder ���g�p����SQL�ڑ��̐ݒ��ۑ�����
$ConnectionString = New-Object -TypeName System.Data.SqlClient.SqlConnectionStringBuilder

#�ڑ�
$ConnectionString['Data Source'] = $datasource #SQLServer�̃z�X�g���w��
$ConnectionString['Initial Catalog'] = $database_name #�f�[�^�x�[�X���w��
$ConnectionString['Integrated Security'] = "TRUE" #Windows�����F�؂𗘗p����ꍇ��"TRUE"

# SQL���̕������ݒ肷��
$SQLQuery = $SQLstring

# DataTable�𗘗p����SQL���s���ʂ��ꎞ�i�[
$resultsDataTable = New-Object System.Data.DataTable

# SQLConnection�ASQLCommand��ݒ肷��
$SqlConnection = New-Object System.Data.SQLClient.SQLConnection($ConnectionString)
$SqlCommand = New-Object System.Data.SQLClient.SQLCommand($SQLQuery, $SqlConnection)

# �f�[�^�x�[�X�֐ڑ�
$SqlConnection.Open()

# ExecuteReader�����s����DataTable�Ƀf�[�^���i�[
$resultsDataTable.Load($SqlCommand.ExecuteReader())

# �f�[�^�x�[�X�ڑ�����
$SqlConnection.Close()

#�ϐ��Ɋi�[
$definition = $resultsDataTable.definition

#2 ��ނ��擾
# SQL���̕������ݒ肷��
$SQLQuery = $SQLstring2

# �i�[����`
$resultsDataTable = New-Object System.Data.DataTable

# SQLConnection�ASQLCommand��ݒ肷��
$SqlConnection = New-Object System.Data.SQLClient.SQLConnection($ConnectionString)
$SqlCommand = New-Object System.Data.SQLClient.SQLCommand($SQLQuery, $SqlConnection)

# �f�[�^�x�[�X�֐ڑ�
$SqlConnection.Open()
$resultsDataTable.Load($SqlCommand.ExecuteReader())
$SqlConnection.Close()

# ���ʂ�ϐ��Ɋi�[
$type = $resultsDataTable.type

#�I�u�W�F�N�g�̎�ނ̕ϊ��p�n�b�V���e�[�u��
$hash = @{
"P "="StoredProcedure";
"TF"="UserDefinedFunction";
"U "="Table";
"V "="View";
"FN"="UserDefinedFunction";
"IF"="UserDefinedFunction";
}

#����p�̕ϐ����`
$datetime = date -Format "g"
$type_def = $hash[$type]

#�o�̓e�L�X�g�쐬
$preamble =@"
USE [$dbname]
GO

/****** Object:  $type_def [$schema].[$name]    Script Date: $datetime ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


"@

$postscript = @"

GO


"@

#�������A��
$output = $preamble + $definition + $postscript

#�o��
echo $output