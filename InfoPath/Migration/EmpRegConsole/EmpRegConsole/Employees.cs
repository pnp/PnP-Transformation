using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmpRegConsole
{
    class Employees
    {
        private string name;
        private string number;
        private string designation;
        private string locaiton;
        private string skills;
        private string userID;
        private string manager;

        public string Manager
        {
            get { return manager; }
            set { manager = value; }
        }

        public string UserID
        {
            get { return userID; }
            set { userID = value; }
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public string Number
        {
            get { return number; }
            set { number = value; }
        }

        public string Designation
        {
            get { return designation; }
            set { designation = value; }
        }
        
        public string Locaiton
        {
            get { return locaiton; }
            set { locaiton = value; }
        }
        
        public string Skills
        {
            get { return skills; }
            set { skills = value; }
        }


    }
}
