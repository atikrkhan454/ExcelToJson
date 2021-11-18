public class Student {
    String Name;
    int Age;
    Long TotalMarks;
    public Student()
    {
        super();
    }
    public Student(String name, int age, int totalMarks)
    {
        super();
        Name = name;
        Age = age;
        TotalMarks = (long) totalMarks;
    }
    public String getName() {
        return Name;
    }
    public void setName(String name) {
        Name = name;
    }
    public int getAge() {
        return Age;
    }
    public void setAge(int age) {
        Age = age;
    }
    public Long getTotalMarks() {
        return TotalMarks;
    }
    public void setTotalMarks(int totalMarks) {
        TotalMarks = (long) totalMarks;
    }
    @Override
    public String toString()
    {
        return "Student [Name=" + Name + ", Age=" + Age + ", TotalMarks=" + TotalMarks + "]";

    }

}
