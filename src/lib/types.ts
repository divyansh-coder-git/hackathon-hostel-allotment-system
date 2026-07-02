export interface Room {
    hostelName: string;
    roomNumber: string;
    capacity: number;
    hostelType: "Boys" | "Girls"| string;
}

export interface Participant {
    id?: number;
    name: string;
    gender: "Male" | "Female" | string;
}

export interface Allocation {
    id: number;
    participantName: string;
    gender: string;
    allottedHostel: string;
    allottedRoom: string;
}
