import { PrismaClient } from '@prisma/client'
import { parse } from 'csv-parse/sync'
import { readFileSync } from 'fs'
import { join } from 'path'

console.log('DATABASE_URL:', process.env.DATABASE_URL)

const prisma = new PrismaClient()

// Function to parse the specific date format in the CSV
function parseDate(dateString: string): string {
  if (!dateString || dateString.trim() === '') {
    return ''
  }
  
  // Handle format like: "Mon Apr 27 2026 00:00:00 GMT+0530 (India Standard Time)"
  try {
    // Extract the main date part before the timezone info
    const mainPart = dateString.split(' GMT')[0]
    const date = new Date(mainPart)
    return date.toISOString().split('T')[0] // Return YYYY-MM-DD format
  } catch (error) {
    console.warn(`Could not parse date: ${dateString}`, error)
    return ''
  }
}

async function main() {
  console.log('Starting seed...')

  // Clear existing data
  await prisma.complaint.deleteMany()
  await prisma.policeStation.deleteMany()
  await prisma.office.deleteMany()
  await prisma.district.deleteMany()
  console.log('Cleared existing data')

  // Load Districts
  const districtData = parse(readFileSync(join(process.cwd(), 'data', 'districts.csv')), {
    columns: true,
    skip_empty_lines: true
  })
  
  const districts = await prisma.district.createMany({
    data: districtData.map((d: any) => ({
      id: Number(d.id),
      name: d.name
    }))
  })
  console.log(`Seeded ${districts.count} districts`)

  // Load Offices
  const officeData = parse(readFileSync(join(process.cwd(), 'data', 'offices.csv')), {
    columns: true,
    skip_empty_lines: true
  })
  
  const offices = await prisma.office.createMany({
    data: officeData.map((o: any) => ({
      id: Number(o.id),
      name: o.name
    }))
  })
  console.log(`Seeded ${offices.count} offices`)

  // Load Police Stations
  const policeStationData = parse(readFileSync(join(process.cwd(), 'data', 'police-stations.csv')), {
    columns: true,
    skip_empty_lines: true
  })
  
  const policeStations = await prisma.policeStation.createMany({
    data: policeStationData.map((ps: any) => ({
      id: Number(ps.id),
      districtId: Number(ps.districtId),
      districtName: ps.districtName,
      name: ps.name
    }))
  })
  console.log(`Seeded ${policeStations.count} police stations`)

   // Load Complaints
   const complaintData = parse(readFileSync(join(process.cwd(), 'data', 'complaints1.csv')), {
     columns: true,
     skip_empty_lines: true
   })
  
  const complaints = await prisma.complaint.createMany({
     data: complaintData
       .filter((c: any) => {
         // Validate required fields
         return c.complRegNum && 
                c.complRegDt && 
                c.districtName && 
                c.districtMasterId !== '' && 
                c.policeStationMasterId !== '' && 
                c.officeMasterId !== '';
       })
       .map((c: any) => ({
      complRegNum: c.complRegNum,
       complRegDt: parseDate(c.complRegDt) || new Date().toISOString().split('T')[0],
      districtName: c.districtName,
       districtMasterId: Number(c.districtMasterId),
       policestationmasterId: Number(c.policeStationMasterId),
       officemasterId: Number(c.officeMasterId),
      complDesc: c.complDesc === '' ? null : c.complDesc,
      complSrno: c.complSrno ? Number(c.complSrno) : null,
      firstName: c.firstName === '' ? null : c.firstName,
      lastName: c.lastName === '' ? null : c.lastName,
      mobile: c.mobile === '' ? null : c.mobile,
      gender: c.gender === '' ? null : c.gender,
      age: c.age ? Number(c.age) : null,
      addressLine1: c.addressLine1 === '' ? null : c.addressLine1,
      addressLine2: c.addressLine2 === '' ? null : c.addressLine2,
      addressLine3: c.addressLine3 === '' ? null : c.addressLine3,
      village: c.village === '' ? null : c.village,
      tehsil: c.tehsil === '' ? null : c.tehsil,
      addressDistrict: c.addressDistrict === '' ? null : c.addressDistrict,
      addressPs: c.addressPs === '' ? null : c.addressPs,
      receptionMode: c.receptionMode === '' ? null : c.receptionMode,
      incidentType: c.incidentType === '' ? null : c.incidentType,
      incidentPlc: c.incidentPlc === '' ? null : c.incidentPlc,
      incidentFromDt: parseDate(c.incidentFromDt),
      incidentToDt: parseDate(c.incidentToDt),
      submitPsCd: c.submitPsCd === '' ? null : c.submitPsCd,
      submitOfficeCd: c.submitOfficeCd === '' ? null : c.submitOfficeCd,
      email: c.email === '' ? null : c.email,
      statusRaw: c.statusRaw === '' ? null : c.statusRaw,
      statusGroup: c.statusGroup === '' ? null : c.statusGroup,
      statusOfComplaint: c.statusOfComplaint === '' ? null : c.statusOfComplaint,
      disposalDate: parseDate(c.disposalDate),
      classOfIncident: c.classOfIncident === '' ? null : c.classOfIncident,
      complaintSource: c.complaintSource === '' ? null : c.complaintSource,
      typeofComplaint: c.typeOfComplaint === '' ? null : c.typeOfComplaint,
      crimeCategory: c.crimeCategory === '' ? null : c.crimeCategory,
      complainantType: c.complainantType === '' ? null : c.complainantType,
      complaintPurpose: c.complaintPurpose === '' ? null : c.complaintPurpose,
      ioDetails: c.ioDetails === '' ? null : c.ioDetails,
      respondentCategories: c.respondentCategories === '' ? null : c.respondentCategories,
      transferDistrictCd: c.transferDistrictCd === '' ? null : c.transferDistrictCd,
      transferOfficeCd: c.transferOfficeCd === '' ? null : c.transferOfficeCd,
      transferPsCd: c.transferPsCd === '' ? null : c.transferPsCd
    }))
  })
  console.log(`Seeded ${complaints.count} complaints`)

  console.log('Seeding completed successfully')
}

main()
  .catch((e) => {
    console.error(e)
    process.exit(1)
  })
  .finally(async () => {
    await prisma.$disconnect()
  })