import React from 'react';

export default function CV() {
  return (
    <div className="min-h-screen bg-gray-100 flex items-center justify-center p-6">
      <div className="bg-white max-w-4xl w-full rounded-2xl shadow-xl p-8">
        {/* Header */}
        <div className="flex items-center justify-between border-b pb-4 mb-6">
          <div>
            <h1 className="text-3xl font-bold text-gray-800">Phan Chitra</h1>
            <p className="text-gray-600">
              Telecommunication & Networking Engineer | Digital Marketing | Event Planner
            </p>
          </div>
          <div className="text-right text-sm">
            <p className="text-gray-700">📞 +855 66 306 250</p>
            <p className="text-gray-700">✉️ phanchitra143@gmail.com</p>
            <p className="text-gray-700">📍 Phnom Penh, Cambodia</p>
          </div>
        </div>

        {/* Summary */}
        <section className="mb-6">
          <h2 className="text-xl font-semibold text-gray-800 mb-2">Summary</h2>
          <p className="text-gray-700">
            5th-year student in Telecommunication and Network Engineering at ITC.
            Experienced in technical services, digital marketing, and event planning.
            Vice President of SANITC 30, passionate about technology, teamwork, and community involvement.
          </p>
        </section>

        {/* Experience */}
        <section className="mb-6">
          <h2 className="text-xl font-semibold text-gray-800 mb-2">Work Experience</h2>
          <div className="space-y-3">
            <div>
              <h3 className="font-semibold text-gray-700">Intersys Solutions Co., Ltd</h3>
              <p className="text-sm text-gray-600">Trainee | Feb – Jul 2025 &amp; Jul – Oct 2024</p>
              <p className="text-gray-700">
                Hands-on experience in technical service operations, ELV systems, and industry best practices.
              </p>
            </div>
            <div>
              <h3 className="font-semibold text-gray-700">Ezecom Company Ltd.</h3>
              <p className="text-sm text-gray-600">Trainee | 2023</p>
              <p className="text-gray-700">
                Worked in Technical Service and Digital Marketing departments. Learned telecommunications operations and
                digital campaigns.
              </p>
            </div>
            <div>
              <h3 className="font-semibold text-gray-700">Spai</h3>
              <p className="text-sm text-gray-600">Trainee | 2023</p>
              <p className="text-gray-700">
                Assisted in digital marketing strategies, social media management, and online promotions.
              </p>
            </div>
          </div>
        </section>

        {/* Education */}
        <section className="mb-6">
          <h2 className="text-xl font-semibold text-gray-800 mb-2">Education</h2>
          <ul className="list-disc list-inside text-gray-700">
            <li>
              <b>2020 – 2025:</b> ITC – Telecommunication &amp; Network Engineering
            </li>
            <li>
              <b>2014 – 2020:</b> Sok An Tram Khnar High School
            </li>
            <li>
              <b>2008 – 2014:</b> Trapeang Kuch Primary School
            </li>
          </ul>
        </section>

        {/* Leadership & Community */}
        <section className="mb-6">
          <h2 className="text-xl font-semibold text-gray-800 mb-2">Volunteer &amp; Leadership</h2>
          <ul className="list-disc list-inside text-gray-700">
            <li>Vice President – SANITC 30</li>
            <li>Organizer – Sangkranta Techno</li>
            <li>Sponsor &amp; Marketing Lead – Sangkranta Techno</li>
            <li>Volunteer – Angkor Songkran, 32nd SEA Games, ASEAN Para Games, CAFEO-40</li>
          </ul>
        </section>

        {/* Skills */}
        <section className="mb-6">
          <h2 className="text-xl font-semibold text-gray-800 mb-2">Skills</h2>
          <div className="grid grid-cols-2 gap-2 text-gray-700">
            <p>Huawei ICT</p>
            <p>Microsoft Office</p>
            <p>Java, C/C++</p>
            <p>Canva, Video Editing</p>
            <p>Leadership &amp; Communication</p>
            <p>Teamwork &amp; Time Management</p>
          </div>
        </section>

        {/* Languages */}
        <section className="mb-6">
          <h2 className="text-xl font-semibold text-gray-800 mb-2">Languages</h2>
          <p className="text-gray-700">Khmer (Native), English (Professional)</p>
        </section>

        {/* Hobbies */}
        <section className="mb-6">
          <h2 className="text-xl font-semibold text-gray-800 mb-2">Hobbies &amp; Interests</h2>
          <p className="text-gray-700">Football, Badminton, Volunteering, Gaming, Traveling, Photo/Video Editing</p>
        </section>

        {/* Reference */}
        <section>
          <h2 className="text-xl font-semibold text-gray-800 mb-2">Reference</h2>
          <p className="text-gray-700">
            Dr. Sreng Sokchenda – Head of Department, ITC <br />
            📞 012 407 910 | ✉️ sokchenda@itc.edu.kh
          </p>
        </section>
      </div>
    </div>
  );
}
