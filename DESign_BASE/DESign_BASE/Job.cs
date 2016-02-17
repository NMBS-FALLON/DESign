using System.Collections.Generic;

namespace DESign_BASE
{
    public class Job
    {
        public string Name { get; set; }
        public string Number { get; set; }
        public string Location { get; set; }
        public List<Joist> Joists { get; set; }
        public List<Girder> Girders { get; set; }
    }


}


