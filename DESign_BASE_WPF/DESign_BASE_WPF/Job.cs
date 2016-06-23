using System.Collections.Generic;

namespace DESign_BASE_WPF
{
    public class Job
    {
        public string Name { get; set; }
        public string Number { get; set; }
        public string Location { get; set; }
        public List<Joist> Joists { get; set; }
        public List<Girder> Girders { get; set; }
        private List<int> listOfStrippedMarks = new List<int>();
        public List<int> ListOfStrippedMarks
        {
            get
            {
                listOfStrippedMarks = CreateListOfStrippedMarks(Girders, Joists);
                return listOfStrippedMarks;
            }
        }

        public static List<int> CreateListOfStrippedMarks(List<Girder> girders, List<Joist> joists)
        {
            List<int> listOfStripedMarks = new List<int>();
            foreach(Girder girder in girders)
            {
                listOfStripedMarks.Add(girder.StrippedNumber);
            }
            foreach(Joist joist in joists)
            {
                listOfStripedMarks.Add(joist.StrippedNumber);
            }
            listOfStripedMarks.Sort();

            return listOfStripedMarks;
        }
    }


}


