import unittest

import numpy as np

from rag_engine import DocumentChunk, _split_words, normalize_rows, retrieve


class RagEngineTests(unittest.TestCase):
    def test_split_words_has_overlap(self):
        words = [f"w{i}" for i in range(12)]
        chunks = _split_words(" ".join(words), max_words=6, overlap_words=2)
        self.assertEqual(chunks[0].split()[-2:], chunks[1].split()[:2])

    def test_normalize_rows_handles_zero_vector(self):
        result = normalize_rows(np.array([[0.0, 0.0], [3.0, 4.0]]))
        self.assertTrue(np.allclose(result[0], [0.0, 0.0]))
        self.assertTrue(np.allclose(result[1], [0.6, 0.8]))

    def test_retrieve_orders_by_cosine_similarity(self):
        chunks = [
            DocumentChunk("1", "Returns", "return rules"),
            DocumentChunk("2", "Delivery", "delivery rules"),
        ]
        vectors = np.array([[1.0, 0.0], [0.0, 1.0]])
        results = retrieve(np.array([[0.9, 0.1]]), vectors, chunks, top_k=2)
        self.assertEqual(results[0][0].section, "Returns")
        self.assertGreater(results[0][1], results[1][1])


if __name__ == "__main__":
    unittest.main()
