def strip_comments(string, markers):
    lines = string.split("\n")

    out = []

    for line in lines:
        for marker in markers:
            if marker in line:
                line = line[: line.index(marker)]
                if line[-1] == " ":
                    line = line[: len(line) - 2]

        out.append(line)

    return "\n".join(out)


print(strip_comments(" apples, pears # and bananas\ngrapes\nbananas !apples", "#!"))
