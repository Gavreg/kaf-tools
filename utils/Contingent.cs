namespace utils
{
    public class Student
    {
        public string Name { set; get; } 
        public string Surname { set; get; }
        public string Lastname { set; get; }
        public bool IsHeadman { set; get; } = false;

       // public Group Group { set; get; }
    }

    public class Group
    {
        public string Name { set; get; }
        public string StandardName { set; get; }
        public string Direction { set; get; }
        public string StandardDirection { set; get; }
        public List<Student> Students { set; get; } = new List<Student>();

        



    }
    public class GroupList
    {
        public List<Group> Groups { set; get; } = new List<Group>();
    }
}