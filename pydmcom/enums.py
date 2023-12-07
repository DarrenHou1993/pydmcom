from enum import Enum, unique


@unique
class KeyCode(Enum):
    N0 = 48
    N1 = 49
    N2 = 50
    N3 = 51
    N4 = 52
    N5 = 53
    N6 = 54
    N7 = 55
    N8 = 56
    N9 = 57
    A = 65
    B = 66
    C = 67
    D = 68
    E = 69
    F = 70
    G = 71
    H = 72
    I = 73
    J = 74
    K = 75
    L = 76
    M = 77
    N = 78
    O = 79
    P = 80
    Q = 81
    R = 82
    S = 83
    T = 84
    U = 85
    V = 86
    W = 87
    X = 88
    Y = 89
    Z = 90

    F1 = 112
    F2 = 113
    F3 = 114
    F4 = 115
    F5 = 116
    F6 = 117
    F7 = 118
    F8 = 119
    F9 = 120
    F10 = 121
    F11 = 122
    F12 = 123

    ENTER = 108
    LEFT_ALT = 18
    LEFT_WIN = 91
    ESC = 27
    CTRL = 17


@unique
class FindDir(Enum):
    LeftToRightAndTopToBottom = 0
    LeftToRightAndBottomToTop = 1
    RightToLeftAndTopToBottom = 2
    RightToLeftAndBottomToTop = 3
    CenterToOutSide = 4
    TopToBottomAndLeftToRight = 5
    TopToBottomAndRightToLeft = 6
    BottomToTopAndLeftToRight = 7
    BottomToTopAndRightToLeft = 8
