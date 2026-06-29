export interface Room {
    hostelName: string;
    roomNumber: string;
    capacity: number;
    hostelType: "Boys" | "Girls"| string;
}

export interface Participant {
    name: string;
    gender: "Male" | "Female" | string;
}

export interface Allocation {
    participantName: string;
    gender: string;
    allottedHostel: string;
    allottedRoom: string;
}