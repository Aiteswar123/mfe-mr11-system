import { Request, Response } from 'express';
import { PrismaClient, RoleCode } from '@prisma/client';
import { processAtomicWorkbookUpload } from './department.service';
import fs from 'fs';

const prisma = new PrismaClient();

const DEPT_NAMES: Record<string, string> = {
  BD: 'Business Development',
  FINANCE: 'Finance',
  SHELLPLAN: 'Shellplan',
  DESIGN: 'Design',
  PLANNING: 'Planning',
  PRODUCTION: 'Production',
  DISPATCH: 'Dispatch',
};

// GET /api/departments/:code
export async function getActiveDepartmentWorkbook(req: Request, res: Response) {
  try {
    const rawParam = (req.params.code || req.params.id || req.params.deptCode || '').trim();
    const deptCode = rawParam.toUpperCase() as RoleCode;

    const validRoleCodes = Object.keys(DEPT_NAMES) as RoleCode[];

    let dept = await prisma.department.findFirst({
      where: {
        OR: [
          validRoleCodes.includes(deptCode) ? { code: deptCode } : undefined,
          { id: rawParam },
        ].filter(Boolean) as any,
      },
      include: {
        activeVersion: {
          include: { uploadedBy: { select: { fullName: true, email: true } } },
        },
      },
    });

    if (!dept && validRoleCodes.includes(deptCode)) {
      dept = await prisma.department.upsert({
        where: { code: deptCode },
        update: {},
        create: {
          code: deptCode,
          name: DEPT_NAMES[deptCode] || deptCode,
          description: `${DEPT_NAMES[deptCode] || deptCode} Department Workbook`,
        },
        include: {
          activeVersion: {
            include: { uploadedBy: { select: { fullName: true, email: true } } },
          },
        },
      });
    }

    if (!dept) {
      return res.status(404).json({
        success: false,
        error: { code: 'DEPT_NOT_FOUND', message: `Department '${rawParam}' not found` },
      });
    }

    return res.json({ success: true, data: dept });
  } catch (err: any) {
    console.error('getActiveDepartmentWorkbook error:', err);
    return res.status(500).json({ success: false, error: { message: err.message } });
  }
}

// POST /api/departments/:code/upload
export async function uploadDepartmentWorkbook(req: Request, res: Response) {
  const rawParam = (req.params.code || req.params.id || '').trim();
  const deptCode = rawParam.toUpperCase() as RoleCode;
  const user = (req as any).user;
  const file = req.file;

  if (!file) {
    return res.status(400).json({
      success: false,
      error: { code: 'FILE_REQUIRED', message: 'No file uploaded' },
    });
  }

  if (user && !user.roles.includes(RoleCode.ADMIN) && !user.roles.includes(deptCode)) {
    if (fs.existsSync(file.path)) fs.unlinkSync(file.path);
    return res.status(403).json({ success: false, error: { code: 'UNAUTHORIZED_DEPARTMENT_UPLOAD' } });
  }

  try {
    await processAtomicWorkbookUpload(
      prisma,
      deptCode,
      file.path,
      file.originalname,
      file.mimetype,
      file.size,
      user?.id
    );

    return res.json({
      success: true,
      message: `${deptCode} workbook uploaded and set as active. Master MR11 regenerated.`,
    });
  } catch (err: any) {
    console.error('uploadDepartmentWorkbook error:', err);
    return res.status(500).json({
      success: false,
      error: { code: 'UPLOAD_FAILED', message: err.message },
    });
  }
}

// GET /api/departments/:code/download
export async function downloadOriginalFile(req: Request, res: Response) {
  try {
    const rawParam = (req.params.code || req.params.id || '').trim();
    const deptCode = rawParam.toUpperCase() as RoleCode;

    const dept = await prisma.department.findFirst({
      where: {
        OR: [{ code: deptCode }, { id: rawParam }],
      },
      include: { activeVersion: true },
    });

    if (!dept?.activeVersion?.storageKey || !fs.existsSync(dept.activeVersion.storageKey)) {
      return res.status(404).json({ success: false, error: { code: 'FILE_NOT_FOUND' } });
    }

    res.setHeader(
      'Content-Type',
      dept.activeVersion.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.download(dept.activeVersion.storageKey, dept.activeVersion.originalFilename);
  } catch (err: any) {
    return res.status(500).json({ success: false, error: { message: err.message } });
  }
}

// GET /api/departments
export async function listDepartments(req: Request, res: Response) {
  try {
    const departments = await prisma.department.findMany({
      orderBy: { code: 'asc' },
      include: {
        activeVersion: {
          select: {
            id: true,
            originalFilename: true,
            status: true,
            uploadedAt: true,
          },
        },
      },
    });
    return res.json({ success: true, data: departments });
  } catch (err: any) {
    return res.status(500).json({ success: false, error: { message: err.message } });
  }
}