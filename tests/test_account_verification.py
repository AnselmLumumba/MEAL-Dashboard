import unittest

from account_verification import (
    create_activation_code,
    hash_activation_code,
    verify_activation_code,
)


class AccountVerificationTests(unittest.TestCase):
    def test_generated_code_is_six_digits_and_hash_verifies(self):
        code = create_activation_code()
        saved_hash = hash_activation_code(code)
        wrong_code = "000001" if code != "000001" else "000002"

        self.assertRegex(code, r"^[0-9]{6}$")
        self.assertNotIn(code, saved_hash)
        self.assertTrue(verify_activation_code(code, saved_hash))
        self.assertFalse(verify_activation_code(wrong_code, saved_hash))

    def test_rejects_malformed_codes_and_hashes(self):
        saved_hash = hash_activation_code("123456")

        self.assertFalse(verify_activation_code("12345", saved_hash))
        self.assertFalse(verify_activation_code("12a456", saved_hash))
        self.assertFalse(verify_activation_code("123456", "invalid"))

    def test_hash_rejects_malformed_codes(self):
        with self.assertRaises(ValueError):
            hash_activation_code("12345")


if __name__ == "__main__":
    unittest.main()
