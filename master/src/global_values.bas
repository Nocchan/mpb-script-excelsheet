Attribute VB_Name = "global_values"
Option Explicit

' �V�[�g���Q�Ƃ�����X�N���v�g���ł��ύX����邱�Ƃ̂Ȃ��萔���`

' Google�h���C�u�̃p�X
Public MPB_WORK_DIRECTORY_PATH As String

' ���[�J�����m���̃p�X
Public LOCAL_WORK_DIRECTORY_PATH As String

' ���c���́����c��
Public DICT_TEAM_NAME As New Dictionary

' ��b�X�y��
Public BASE_ACCIDENT_RATE As Single

' ���c���́�����a�@�K�p��
Public DICT_ACCIDENT_HDCP As New Dictionary

' �X�y����l���W��
Public DICT_ACCIDENT_COEFFICIENT As New Dictionary

' �\�X�y���I�p����
Public DICT_ACCIDENT_LENGTH_RATE As New Dictionary

' ���X�y���I�p�񎟌�����
Public DICT_ACCIDENT_MARGIN_DICT As New Dictionary
Public DICT_ACCIDENT_MARGIN_S As New Dictionary
Public DICT_ACCIDENT_MARGIN_A As New Dictionary
Public DICT_ACCIDENT_MARGIN_B As New Dictionary
Public DICT_ACCIDENT_MARGIN_C As New Dictionary
Public DICT_ACCIDENT_MARGIN_D As New Dictionary
Public DICT_ACCIDENT_MARGIN_E As New Dictionary
Public DICT_ACCIDENT_MARGIN_F As New Dictionary
Public DICT_ACCIDENT_MARGIN_G As New Dictionary

' ����X�y���e���I�p�񎟌�����
' MPB�j���[�X�o�͂� ���`�[�������I�薼�I��XXX �Ƃ���ꍇ�� XXX ���`
Public DICT_ACCIDENT_INFORMATION_PITCHER_DICT As New Dictionary
Public DICT_ACCIDENT_INFORMATION_PITCHER_1 As New Dictionary
Public DICT_ACCIDENT_INFORMATION_PITCHER_2 As New Dictionary
Public DICT_ACCIDENT_INFORMATION_PITCHER_5 As New Dictionary
Public DICT_ACCIDENT_INFORMATION_PITCHER_8 As New Dictionary
Public DICT_ACCIDENT_INFORMATION_PITCHER_24 As New Dictionary

' ���X�y���e���I�p�񎟌�����
' MPB�j���[�X�o�͂� ���`�[�������I�薼�I��XXX �Ƃ���ꍇ�� XXX ���`
Public DICT_ACCIDENT_INFORMATION_FIELDER_DICT As New Dictionary
Public DICT_ACCIDENT_INFORMATION_FIELDER_1 As New Dictionary
Public DICT_ACCIDENT_INFORMATION_FIELDER_2 As New Dictionary
Public DICT_ACCIDENT_INFORMATION_FIELDER_5 As New Dictionary
Public DICT_ACCIDENT_INFORMATION_FIELDER_8 As New Dictionary
Public DICT_ACCIDENT_INFORMATION_FIELDER_24 As New Dictionary

Public Function Definition()

    MPB_WORK_DIRECTORY_PATH = "C:\Users\TaiNo\�}�C�h���C�u\MPB\1-�܂�"
    LOCAL_WORK_DIRECTORY_PATH = "C:\Users\TaiNo\Desktop\MPB\1-�܂�"

    With DICT_TEAM_NAME
        .Add "G", "�W���C�A���c"
        .Add "M", "�}���[���Y"
        .Add "T", "�^�C�K�[�X"
        .Add "L", "���C�I���Y"
        .Add "E", "�C�[�O���X"
    End With

    BASE_ACCIDENT_RATE = 0.01

    With DICT_ACCIDENT_HDCP
        .Add "G", 1#
        .Add "M", 1#
        .Add "T", 1#
        .Add "L", 1#
        .Add "E", 1#
    End With

    With DICT_ACCIDENT_COEFFICIENT
        .Add "S", 0.01
        .Add "A", 0.3
        .Add "B", 0.5
        .Add "C", 0.8
        .Add "D", 1#
        .Add "E", 1.2
        .Add "F", 2#
        .Add "G", 4#
        .Add "n", 0#
    End With

    With DICT_ACCIDENT_LENGTH_RATE
        .Add 1, 36#
        .Add 2, 40#
        .Add 5, 12#
        .Add 8, 8#
        .Add 24, 4#
    End With

    With DICT_ACCIDENT_MARGIN_DICT
        .Add "S", DICT_ACCIDENT_MARGIN_S
        .Add "A", DICT_ACCIDENT_MARGIN_A
        .Add "B", DICT_ACCIDENT_MARGIN_B
        .Add "C", DICT_ACCIDENT_MARGIN_C
        .Add "D", DICT_ACCIDENT_MARGIN_D
        .Add "E", DICT_ACCIDENT_MARGIN_E
        .Add "F", DICT_ACCIDENT_MARGIN_F
        .Add "G", DICT_ACCIDENT_MARGIN_G
    End With

    With DICT_ACCIDENT_MARGIN_S
        .Add -1, 30#
        .Add 0, 70#
    End With

    With DICT_ACCIDENT_MARGIN_A
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 30#
    End With

    With DICT_ACCIDENT_MARGIN_B
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 20#
        .Add 2, 10#
    End With

    With DICT_ACCIDENT_MARGIN_C
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 15#
        .Add 2, 10#
        .Add 3, 5#
    End With

    With DICT_ACCIDENT_MARGIN_D
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 12#
        .Add 2, 9#
        .Add 3, 6#
        .Add 4, 3#
    End With

    With DICT_ACCIDENT_MARGIN_E
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 10#
        .Add 2, 8#
        .Add 3, 6#
        .Add 4, 4#
        .Add 5, 2#
    End With

    With DICT_ACCIDENT_MARGIN_F
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 8.57
        .Add 2, 7.14
        .Add 3, 5.71
        .Add 4, 4.29
        .Add 5, 2.86
        .Add 6, 1.43
    End With

    With DICT_ACCIDENT_MARGIN_G
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 7.5
        .Add 2, 6.43
        .Add 3, 5.36
        .Add 4, 4.29
        .Add 5, 3.21
        .Add 6, 2.14
        .Add 7, 1.07
    End With

    With DICT_ACCIDENT_INFORMATION_PITCHER_DICT
        .Add 1, DICT_ACCIDENT_INFORMATION_PITCHER_1
        .Add 2, DICT_ACCIDENT_INFORMATION_PITCHER_2
        .Add 5, DICT_ACCIDENT_INFORMATION_PITCHER_5
        .Add 8, DICT_ACCIDENT_INFORMATION_PITCHER_8
        .Add 24, DICT_ACCIDENT_INFORMATION_PITCHER_24
    End With

    With DICT_ACCIDENT_INFORMATION_PITCHER_1
        .Add "�I�肪�A�Ē����̂��߁A���߂͈�x�x���`����O���Ƃ̂��Ƃł��B", 1
        .Add "�I��́A�㎈�̃R���f�B�V�����s�ǂɂ��A�O�̂��ߎ��߂̃x���`����O���Ƃ̂��Ƃł��B", 1
        .Add "�I��́A�����̃R���f�B�V�����s�ǂɂ��A�O�̂��ߎ��߂̃x���`����O���Ƃ̂��Ƃł��B", 1
        .Add "�I��́A���K���ɕI�̈�a����i�������߁A�O�̂��ߎ��߂̃x���`����O���Ƃ̂��Ƃł��B", 1
        .Add "�I��́A���K���Ɍ��̈�a����i�������߁A�O�̂��ߎ��߂̃x���`����O���Ƃ̂��Ƃł��B", 1
        .Add "�I��́A���K���ɍ��̈�a����i�������߁A�O�̂��ߎ��߂̃x���`����O���Ƃ̂��Ƃł��B", 1
        .Add "�I��ɁA�s�b�`���[���C�i�[���󂯂�A�N�V�f���g�B�O�̂��ߎ��߂̓x���`����O���Ƃ̂��Ƃł��B", 1
        .Add "�I��́A�w�ɂł����}���̉e���ŁA�O�̂��ߎ��߂̃x���`����O���Ƃ̂��Ƃł��B", 1
        .Add "�I��́A�̒��s�ǂ̂��߁A�R���i����œo�^��������܂����B", 1
    End With

    With DICT_ACCIDENT_INFORMATION_PITCHER_2
        .Add "�I�肪�㎈�̃R���f�B�V�����s�ǂƂ̂��ƁB�����͍s�킸�A�l�q��������j�ł��B", 1
        .Add "�I�肪�����̃R���f�B�V�����s�ǂƂ̂��ƁB�����͍s�킸�A�l�q��������j�ł��B", 1
        .Add "�I�肪�I�̈�a����i�����Ƃ̂��ƁB�����͍s�킸�A�l�q��������j�ł��B", 1
        .Add "�I�肪���̈�a����i�����Ƃ̂��ƁB�����͍s�킸�A�l�q��������j�ł��B", 1
        .Add "�I�肪���̈�a����i�����Ƃ̂��ƁB�����͍s�킸�A�l�q��������j�ł��B", 1
        .Add "�I��̎w�Ƀ}�����ł����Ƃ̂��ƁB�����͍s�킸�A�l�q��������j�ł��B", 1
    End With

    With DICT_ACCIDENT_INFORMATION_PITCHER_5
        .Add "�I�肪�㎈�̃R���f�B�V�����s�ǂƂ̂��ƁB��x�������A�l�q��������j�ł��B", 1
        .Add "�I�肪�����̃R���f�B�V�����s�ǂƂ̂��ƁB��x�������A�l�q��������j�ł��B", 1
        .Add "�I�肪�I�̈�a����i�����Ƃ̂��ƁB��x�������A�l�q��������j�ł��B", 1
        .Add "�I�肪���̈�a����i�����Ƃ̂��ƁB��x�������A�l�q��������j�ł��B", 1
        .Add "�I�肪���̈�a����i�����Ƃ̂��ƁB��x�������A�l�q��������j�ł��B", 1
    End With

    With DICT_ACCIDENT_INFORMATION_PITCHER_8
        .Add "�I��́A�㎈�̃R���f�B�V�����s�ǂ̂��߁A�o�^�������A���Âɐ�O����Ƃ̂��Ƃł��B", 1
        .Add "�I��́A�����̃R���f�B�V�����s�ǂ̂��߁A�o�^�������A���Âɐ�O����Ƃ̂��Ƃł��B", 1
        .Add "�I��́A�I�̈�a����i�������߁A�o�^�������A���Âɐ�O����Ƃ̂��Ƃł��B", 1
        .Add "�I��́A���̈�a����i�������߁A�o�^�������A���Âɐ�O����Ƃ̂��Ƃł��B", 1
        .Add "�I��́A���ɂ̂��߁A�o�^�������A���Âɐ�O����Ƃ̂��Ƃł��B", 1
    End With

    With DICT_ACCIDENT_INFORMATION_PITCHER_24
        .Add "�I�肪�A���߂̓o�̍ۂɕI��ɂ߁A�ً}�~�B�ߓ����Ɏ�p���s���Ƃ̂��ƂŁA���V�[�Y�����̕��A�͐�]�I�Ƃ݂��܂��B", 1
        .Add "�I�肪�A���߂̓o�̍ۂɌ���ɂ߁A�ً}�~�B�ߓ����Ɏ�p���s���Ƃ̂��ƂŁA���V�[�Y�����̕��A�͐�]�I�Ƃ݂��܂��B", 1
    End With

    With DICT_ACCIDENT_INFORMATION_FIELDER_DICT
        .Add 1, DICT_ACCIDENT_INFORMATION_FIELDER_1
        .Add 2, DICT_ACCIDENT_INFORMATION_FIELDER_2
        .Add 5, DICT_ACCIDENT_INFORMATION_FIELDER_5
        .Add 8, DICT_ACCIDENT_INFORMATION_FIELDER_8
        .Add 24, DICT_ACCIDENT_INFORMATION_FIELDER_24
    End With

    With DICT_ACCIDENT_INFORMATION_FIELDER_1
        .Add "�I�肪�A�Ē����̂��߁A���߂͈�x�x���`����O���Ƃ̂��Ƃł��B", 1
        .Add "�I��́A�㎈�̃R���f�B�V�����s�ǂɂ��A�O�̂��ߎ��߂̃x���`����O���Ƃ̂��Ƃł��B", 1
        .Add "�I��́A�����̃R���f�B�V�����s�ǂɂ��A�O�̂��ߎ��߂̃x���`����O���Ƃ̂��Ƃł��B", 1
        .Add "�I��́A���K���ɑ������̈�a����i�������߁A�O�̂��ߎ��߂̃x���`����O���Ƃ̂��Ƃł��B", 1
        .Add "�I��́A���K���ɍ��̈�a����i�������߁A�O�̂��ߎ��߂̃x���`����O���Ƃ̂��Ƃł��B", 1
        .Add "�I��́A�����ɂ��Ŗo�̉e���ŁA�O�̂��ߎ��߂̃x���`����O���Ƃ̂��Ƃł��B", 1
        .Add "�I��́A���ے��̃A�N�V�f���g�̉e���ŁA�O�̂��ߎ��߂̃x���`����O���Ƃ̂��Ƃł��B", 1
        .Add "�I��́A�̒��s�ǂ�i�������߁A�R���i����œo�^��������܂����B", 1
    End With

    With DICT_ACCIDENT_INFORMATION_FIELDER_2
        .Add "�I�肪�㎈�̃R���f�B�V�����s�ǂƂ̂��ƁB�����͍s�킸�A�l�q��������j�ł��B", 1
        .Add "�I�肪�����̃R���f�B�V�����s�ǂƂ̂��ƁB�����͍s�킸�A�l�q��������j�ł��B", 1
        .Add "�I�肪�������̈�a����i�����Ƃ̂��ƁB�����͍s�킸�A�l�q��������j�ł��B", 1
        .Add "�I�肪���̈�a����i�����Ƃ̂��ƁB�����͍s�킸�A�l�q��������j�ł��B", 1
        .Add "�I��́A�������󂯎s���̕a�@����f�B�����͍s�킸�A�l�q��������j�ł��B", 1
        .Add "�I�肪�A���ے��̃A�N�V�f���g�œr����サ�܂����B�����͍s�킸�A�l�q��������j�ł��B", 1
    End With

    With DICT_ACCIDENT_INFORMATION_FIELDER_5
        .Add "�I�肪�㎈�̃R���f�B�V�����s�ǂƂ̂��ƁB��x�������A�l�q��������j�ł��B", 1
        .Add "�I�肪�����̃R���f�B�V�����s�ǂƂ̂��ƁB��x�������A�l�q��������j�ł��B", 1
        .Add "�I�肪�������̈�a����i�����Ƃ̂��ƁB��x�������A�l�q��������j�ł��B", 1
        .Add "�I�肪���̈�a����i�����Ƃ̂��ƁB��x�������A�l�q��������j�ł��B", 1
        .Add "�I��́A�������󂯎s���̕a�@����f�B��x�������A�l�q��������j�ł��B", 1
        .Add "�I�肪�A���ے��̃A�N�V�f���g�œr����サ�܂����B��x�������A�l�q��������j�ł��B", 1
    End With

    With DICT_ACCIDENT_INFORMATION_FIELDER_8
        .Add "�I��́A�㎈�̃R���f�B�V�����s�ǂ̂��߁A�o�^�������A���Âɐ�O����Ƃ̂��Ƃł��B", 1
        .Add "�I��́A�����̃R���f�B�V�����s�ǂ̂��߁A�o�^�������A���Âɐ�O����Ƃ̂��Ƃł��B", 1
        .Add "�I��́A���ɂ̂��߁A�o�^�������A���Âɐ�O����Ƃ̂��Ƃł��B", 1
        .Add "�I��́A�������󂯍��܁B�o�^�������A���Âɐ�O����Ƃ̂��Ƃł��B", 1
        .Add "�I��́A���ے��̃A�N�V�f���g�œr�����A������Ɛf�f�B�o�^�������A���Âɐ�O����Ƃ̂��Ƃł��B", 1
    End With

    With DICT_ACCIDENT_INFORMATION_FIELDER_24
        .Add "�I�肪�A����̍ۂɕG��ɂߕa�@�ɒ��s�B�O�\���x�ё����Ɛf�f����܂����B���V�[�Y�����̕��A�͐�]�I�Ƃ݂��܂��B", 1
        .Add "�I��́A���ɂ�i���a�@����f�����Ƃ���A�ŊԔw���j�A�Ɛf�f����܂����B�ߓ����Ɏ�p���s���Ƃ̂��ƂŁA���V�[�Y�����̕��A�͐�]�I�Ƃ݂��܂��B", 1
    End With

End Function
