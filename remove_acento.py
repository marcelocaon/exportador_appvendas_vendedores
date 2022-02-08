def remove_acento(txt):
    return normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')
