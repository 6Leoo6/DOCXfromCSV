ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
N = len(ALPHABET)
ALPHABET_INDEX = {d: i for i, d in enumerate(ALPHABET, 1)}


def int_to_letter(num):
    if num == 0:
        return ""
    else:
        q, r = divmod(num - 1, N)
        return int_to_letter(q) + ALPHABET[r]

if __name__ == '__main__':
    print(int_to_letter(1))