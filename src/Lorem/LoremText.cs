// ---------------------------------------------------------------------------------------------------------------
// Name:     LoremText.cs
// Purpose:  Makes classic Lorem Ipsum filler text. The words all come from the standard Lorem Ipsum passage.
//           The first paragraph of a run opens with the passage's own first sentence; every other sentence is
//           built from the passage's words, 8 to 16 words long, with a comma in the longer ones. Titles are 2
//           to 4 words with each word capitalized.
// Revision: 1
// DTG:      261025030000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Text;

namespace OpenTools
{
    internal sealed class LoremText
    {
        private const string OpeningSentence = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do "
            + "eiusmod tempor incididunt ut labore et dolore magna aliqua.";
        private const int ShortestSentence = 8;
        private const int LongestSentence = 16;
        private const int CommaFrom = 10;
        private const int ShortestTitle = 2;
        private const int LongestTitle = 4;

        private static readonly string[] Words =
        {
            "lorem", "ipsum", "dolor", "sit", "amet", "consectetur", "adipiscing", "elit", "sed", "do", "eiusmod",
            "tempor", "incididunt", "ut", "labore", "et", "dolore", "magna", "aliqua", "enim", "ad", "minim",
            "veniam", "quis", "nostrud", "exercitation", "ullamco", "laboris", "nisi", "aliquip", "ex", "ea",
            "commodo", "consequat", "duis", "aute", "irure", "in", "reprehenderit", "voluptate", "velit", "esse",
            "cillum", "eu", "fugiat", "nulla", "pariatur", "excepteur", "sint", "occaecat", "cupidatat", "non",
            "proident", "sunt", "culpa", "qui", "officia", "deserunt", "mollit", "anim", "id", "est", "laborum"
        };

        private readonly Random _random = new Random();
        private bool _opened;

        // ----- Pieces -----

        public string Paragraph(int sentences)
        {
            StringBuilder text = new StringBuilder();
            for (int i = 0; i < sentences; i++)
            {
                if (i > 0)
                {
                    text.Append(' ');
                }

                text.Append(Sentence());
            }

            return text.ToString();
        }

        public string Sentence()
        {
            if (!_opened)
            {
                _opened = true;
                return OpeningSentence;
            }

            int count = _random.Next(ShortestSentence, LongestSentence + 1);
            int comma = count >= CommaFrom ? _random.Next(3, count - 3) : -1;
            StringBuilder text = new StringBuilder();
            for (int i = 0; i < count; i++)
            {
                string word = Words[_random.Next(Words.Length)];
                text.Append(i == 0 ? Capital(word) : word);
                text.Append(i == count - 1 ? "." : (i == comma ? ", " : " "));
            }

            return text.ToString();
        }

        public string Title()
        {
            int count = _random.Next(ShortestTitle, LongestTitle + 1);
            StringBuilder text = new StringBuilder();
            for (int i = 0; i < count; i++)
            {
                if (i > 0)
                {
                    text.Append(' ');
                }

                text.Append(Capital(Words[_random.Next(Words.Length)]));
            }

            return text.ToString();
        }

        public int Between(int low, int high)
        {
            if (high < low)
            {
                int swap = low;
                low = high;
                high = swap;
            }

            return _random.Next(low, high + 1);
        }

        private static string Capital(string word)
        {
            return char.ToUpperInvariant(word[0]) + word.Substring(1);
        }
    }
}
