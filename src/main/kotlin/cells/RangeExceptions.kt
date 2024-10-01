package org.example.cells

class InvalidAddressException(message: String = ""): RuntimeException(message)

class MismatchedDimensionsException(message: String = "Dimensions are mismatched"): RuntimeException(message)