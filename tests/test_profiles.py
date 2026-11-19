import unittest

from tests import stubs  # noqa: F401

import profiles as prof

A = prof.Profile("Гарнитура", mic="{m1}", out="{o1}", eq_mic="Голос", eq_out="Бас", hotkey="ctrl+alt+1")
B = prof.Profile("Колонки", mic="{m2}", out="{o2}")


class ProfileTests(unittest.TestCase):
    def test_roundtrip(self):
        self.assertEqual(prof.load(prof.dump([A, B])), [A, B])

    def test_missing_eq_means_do_not_touch(self):
        self.assertNotIn("eq_mic", B.to_dict())
        self.assertIsNone(prof.Profile.from_dict({"name": "x"}).eq_out)

    def test_load_skips_broken_and_duplicates(self):
        raw = [A.to_dict(), {"name": ""}, "мусор", {"name": "гарнитура", "mic": "other"}, B.to_dict()]
        self.assertEqual([p.name for p in prof.load(raw)], ["Гарнитура", "Колонки"])
        self.assertEqual(prof.load({"name": "не список"}), [])
        self.assertEqual(prof.load(None), [])

    def test_find_case_insensitive(self):
        self.assertIs(prof.find([A, B], " колонки "), B)
        self.assertIsNone(prof.find([A, B], "нет"))

    def test_next_after_cycles(self):
        self.assertIs(prof.next_after([A, B], "Гарнитура"), B)
        self.assertIs(prof.next_after([A, B], "Колонки"), A)
        self.assertIs(prof.next_after([A, B], ""), A)
        self.assertIsNone(prof.next_after([], "x"))

    def test_upsert_and_remove(self):
        updated = prof.Profile("ГАРНИТУРА", mic="{new}")
        self.assertEqual(prof.upsert([A, B], updated), [updated, B])
        new = prof.Profile("Стрим")
        self.assertEqual(prof.upsert([A, B], new), [A, B, new])
        self.assertEqual(prof.remove([A, B], "гарнитура"), [B])

    def test_auto_name(self):
        self.assertEqual(prof.auto_name("Наушники", "Микрофон"), "Наушники + Микрофон")
        self.assertEqual(prof.auto_name("", "Микрофон"), "Микрофон")
        self.assertEqual(prof.auto_name("", ""), "Профиль")


if __name__ == "__main__":
    unittest.main()
