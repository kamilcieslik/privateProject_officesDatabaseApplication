﻿namespace RogalApplication.Model
{
    public class Customer:Entity
    {
        public string Name { get; set; }
        public string PhoneNumber { get; set; }
        public string Email { get; set; }
        public string Trade { get; set; }
        public string Province { get; set; }
        public string City { get; set; }
        public string PostalCode { get; set; }
        public string Street { get; set; }
        public string Description { get; set; }
        public string Comments { get; set; }
    }
}
