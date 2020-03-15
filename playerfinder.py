for table in list:  # Sorts through tables
    # Sorts through gold count per player in given table
    tableGold = table.find_all("div", {"class": "sb-p-stat sb-p-stat-gold"})
    # Sorts through name per player in given table
    tableName = table.find_all("div", {"class": "sb-p-name"})
    gameTime = table.find_all("tr", id=lambda value: value and value.startswith("sb-allw"))

    print(gameTime)
    for EachPart in gameTime:
        print("alert")

    for i in range(10):  # Max 10 players per match
        if in_list(tableName[i].text, masterArray) != -1:
            value_index = in_list(tableName[i].text, masterArray)
            masterArray[value_index].append(Decimal(tableGold[i].text[:-1]))
        else:
            masterArray.append([tableName[i].text, Decimal(tableGold[i].text[:-1])])