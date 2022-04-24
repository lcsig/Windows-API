Imports System.Security.Cryptography
Imports System.Text

Public Class Hash
    Public Shared Function Compute(ByVal plainText As String) As String
        Return Compute(plainText, Nothing)
    End Function

    Private Shared Function Compute(ByVal plainText As String, _
                                       ByVal saltBytes() As Byte) _
                           As String

        ' If salt is not specified, generate it on the fly.
        If (saltBytes Is Nothing) Then

            ' Generate a random number for the size of the salt.
            Dim random As New Random()

            Dim saltSize As Integer = random.Next(4, 8)

            ' Allocate a byte array, which will hold the salt.
            saltBytes = New Byte(saltSize - 1) {}

            ' Initialize a random number generator.
            Dim rng As New RNGCryptoServiceProvider()

            ' Fill the salt with cryptographically strong byte values.
            rng.GetNonZeroBytes(saltBytes)
        End If

        ' Convert plain text into a byte array.
        Dim plainTextBytes As Byte()
        plainTextBytes = Encoding.UTF8.GetBytes(plainText)

        ' Allocate array, which will hold plain text and salt.
        Dim plainTextWithSaltBytes() As Byte = _
            New Byte(plainTextBytes.Length + saltBytes.Length - 1) {}

        ' Copy plain text bytes into resulting array.
        Dim i As Integer
        For i = 0 To plainTextBytes.Length - 1
            plainTextWithSaltBytes(i) = plainTextBytes(i)
        Next i

        ' Append salt bytes to the resulting array.
        For i = 0 To saltBytes.Length - 1
            plainTextWithSaltBytes(plainTextBytes.Length + i) = saltBytes(i)
        Next i

        ' Because we support multiple hashing algorithms, we must define
        ' hash object as a common (abstract) base class. We will specify the
        ' actual hashing algorithm class later during object creation.
        Dim hash As HashAlgorithm
        hash = New SHA512Managed()

        ' Compute hash value of our plain text with appended salt.
        Dim hashBytes As Byte()
        hashBytes = hash.ComputeHash(plainTextWithSaltBytes)

        ' Create array which will hold hash and original salt bytes.
        Dim hashWithSaltBytes() As Byte = _
                                   New Byte(hashBytes.Length + _
                                            saltBytes.Length - 1) {}

        ' Copy hash bytes into resulting array.
        For i = 0 To hashBytes.Length - 1
            hashWithSaltBytes(i) = hashBytes(i)
        Next i

        ' Append salt bytes to the result.
        For i = 0 To saltBytes.Length - 1
            hashWithSaltBytes(hashBytes.Length + i) = saltBytes(i)
        Next i

        ' Convert result into a base64-encoded string.
        Dim hashValue As String
        hashValue = Convert.ToBase64String(hashWithSaltBytes)

        ' Return the result.
        Compute = hashValue
    End Function

    Public Shared Function Verify(ByVal plainText As String, _
                                      ByVal hashValue As String) _
                           As Boolean

        ' Convert base64-encoded hash value into a byte array.
        Dim hashWithSaltBytes As Byte()
        hashWithSaltBytes = Convert.FromBase64String(hashValue)
        If hashWithSaltBytes.Length = 0 AndAlso String.IsNullOrEmpty(plainText) Then
            'empty password, empty hash
            Return True
        End If

        ' We must know size of hash (without salt).
        Dim hashSizeInBits As Integer
        Dim hashSizeInBytes As Integer

        hashSizeInBits = 512

        ' Convert size of hash from bits to bytes.
        hashSizeInBytes = CInt(hashSizeInBits / 8)

        ' Make sure that the specified hash value is long enough.
        If (hashWithSaltBytes.Length < hashSizeInBytes) Then
            Return False
        End If

        ' Allocate array to hold original salt bytes retrieved from hash.
        Dim saltBytes() As Byte = New Byte(hashWithSaltBytes.Length - _
                                           hashSizeInBytes - 1) {}

        ' Copy salt from the end of the hash to the new array.
        Dim I As Integer
        For I = 0 To saltBytes.Length - 1
            saltBytes(I) = hashWithSaltBytes(hashSizeInBytes + I)
        Next I

        ' Compute a new hash string.
        Dim expectedHashString As String
        expectedHashString = Compute(plainText, saltBytes)

        ' If the computed hash matches the specified hash,
        ' the plain text value must be correct.
        Return (hashValue = expectedHashString)
    End Function
End Class
